#!/bin/bash
LOGIN="LOGIN@gmail.com"
PASS="PASS"
IMAPSERVER="imap.gmail.com"
SMTPSERVER="smtp.gmail.com"

#define some vars
ROOTDIR="Emails_$LOGIN"
TMPDIR="temp"
MERGEDFILE="Merged.pdf"

TIMEOUT=10

#create root dir
mkdir -p "$ROOTDIR"
cd "$ROOTDIR"

#get unseen messages ids
UNSEEN=`curl -s "imaps://$IMAPSERVER/INBOX?UNSEEN" -u "$LOGIN:$PASS"`
#print UNSEEN variable, remove new line character, split by spaces, get columns from 3 to last
UIDS=`echo "$UNSEEN" | tr -d '\r\n' | cut -d' ' -f 3-`

#for every UID
for ITEM in $UIDS; do

    #download full email body
    echo "Downloading message with UID=$ITEM"
    EMAIL=`curl -s --max-time $TIMEOUT "imaps://$IMAPSERVER/INBOX;UID=$ITEM" -u "$LOGIN:$PASS"`
    
    #get filename from its header and save email to it
    EMAILFILE=`echo "$EMAIL" | formail -z -x Date: -x From: | tr -d '\r' | tr ' \n' '_'`
    echo "$EMAIL" > "$EMAILFILE"

    #extract attachments from email to temp dir and parse attachmets names
    mkdir -p "$TMPDIR"
    cd "$TMPDIR"
    ATTACHMENTS=`munpack -f -q "../$EMAILFILE" | cut -d' ' -f 1`

    #for every attachment
    for FILE in $ATTACHMENTS; do

	#if it's pdf or excel
        FILEFORMAT=`file "$FILE"`
        if [[ $FILEFORMAT =~ "PDF" ]] || [[ $FILEFORMAT =~ "Excel" ]]; then
            echo "Merging attachment $FILE"

            #if it's excel, convert to pdf
            if [[ $FILEFORMAT =~ "Excel" ]]; then
                libreoffice --headless --convert-to pdf "$FILE" > /dev/null
                FILE="${FILE%\.xls*}.pdf"
            fi

            #if merged file not exist, just copy current file
            if [ ! -f "../$MERGEDFILE" ]; then
                cp -f "$FILE" "../$MERGEDFILE"
	    #else merge current file with merged.pdf
            else
		mv -f "../$MERGEDFILE" "$MERGEDFILE"
                pdftk "$MERGEDFILE" "$FILE" cat output "../$MERGEDFILE"
            fi 
        fi
    done
    
    cd ..
    #send email back if there some attachments
    if [[ "$ATTACHMENTS" != "Did" ]]; then
        #get sender and receiver from email
        SENDER=`echo "$EMAIL" | formail -x From: | egrep -o "<.*>" | tr -d '<>'`
        RECEIVER=`echo "$EMAIL" | formail -x To: | egrep -o "<.*>" | tr -d '<>'`
        #send message back
        echo "Sending message back to $SENDER"
        curl --ssl "smtp://$SMTPSERVER/${LOGIN/@/_at_}" -u "$LOGIN:$PASS" --mail-from "$RECEIVER" --mail-rcpt "$SENDER" --upload-file "$EMAILFILE" -s
    fi

    #remove temp files
    rm -f -r "$TMPDIR"
done

