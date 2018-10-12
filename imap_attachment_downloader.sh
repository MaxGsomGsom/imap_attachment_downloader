#!/bin/bash

#Install required packages first:
#sudo apt-get install default-jre libreoffice curl procmail mpack pdftk

#=== VARIABLES ===

#mode: 1-download from server, 2-only process existing messages
MODE=1

#login, pass and MESSAGE server
LOGIN="LOGIN@gmail.com"
PASS="PASS"
IMAPSERVER="imap.gmail.com"
SMTPSERVER="smtp.gmail.com"

#dir and files names
ROOTDIR="Messages_$LOGIN"
DOWNLOADEDDIR="$ROOTDIR/Downloaded" #in mode 2 select dir with existing messages
PROCESSEDDIR="$ROOTDIR/Processed" #all messages will be put there after processing
TEMPDIR="$ROOTDIR/Temp" #do not use existing dir
MERGEDFILE="$ROOTDIR/Merged.pdf"

#sometimes CURL does not close connection after downloading big message (>5Mb)
#set this timeout accoding to your connection speed
TIMEOUT=10

#=== DOWNLOADING ===

#create required dirs
mkdir -p "$DOWNLOADEDDIR"
mkdir -p "$PROCESSEDDIR"
mkdir -p "$TEMPDIR"

if [[ $MODE == 1 ]]; then

    echo "=== DOWNLOADING ==="

    #get unseen messages uids
    UNSEEN=`curl -s "imaps://$IMAPSERVER/INBOX?UNSEEN" -u "$LOGIN:$PASS"`
    #process responce from server: remove new line characters, split by spaces, get columns from 3 to last
    UIDS=`echo "$UNSEEN" | tr -d '\r\n' | cut -d' ' -f 3-`
    
    #for every message UID
    for ITEM in $UIDS; do
    
        #download full message body
        echo "Downloading message with UID $ITEM"
        MESSAGE=`curl -s --max-time $TIMEOUT "imaps://$IMAPSERVER/INBOX;UID=$ITEM" -u "$LOGIN:$PASS"`
    
        #create filename from header and save message
        MESSAGEFILE=`echo "$MESSAGE" | formail -z -x Date: -x From: | tr -d '\r' | tr ' \n' '_'`
        echo "$MESSAGE" > "$DOWNLOADEDDIR/$MESSAGEFILE"
    
    done
fi

#=== PROCESSING ===
echo "=== PROCESSING ==="

#for every message UID
for MESSAGEFILE in $DOWNLOADEDDIR/*; do

    #check if it is real message
    MESSAGEFILEFORMAT=`file "$MESSAGEFILE"`
    if [[ ! "$MESSAGEFILEFORMAT" =~ "SMTP" && ! "$MESSAGEFILEFORMAT" =~ "MIME" ]]; then
        continue
    fi

    echo "Processing message $(basename "$MESSAGEFILE")"

    #extract attachments from message and parse attachments names
    ATTACHMENTS=`munpack -f -q "$PWD/$MESSAGEFILE" -C "$TEMPDIR" | cut -d' ' -f 1`

    #for every attachment
    for FILE in $ATTACHMENTS; do

	#if it's pdf or excel
        FILEFORMAT=`file "$TEMPDIR/$FILE"`
        if [[ "$FILEFORMAT" =~ "PDF" || "$FILEFORMAT" =~ "Excel" ]]; then
            echo "Merging attachment $FILE"

            #if it's excel, convert to pdf
            if [[ "$FILEFORMAT" =~ "Excel" ]]; then
                libreoffice --headless --convert-to pdf "$TEMPDIR/$FILE" --outdir "$TEMPDIR" > /dev/null
                FILE="${FILE%\.xls*}.pdf"
            fi

            #if merged file not exist, just copy current file
            if [[ ! -f "$MERGEDFILE" ]]; then
                cp -f "$TEMPDIR/$FILE" "$MERGEDFILE"
	    #else merge current file
            else
		mv -f "$MERGEDFILE" "$TEMPDIR/.merged_temp"
                pdftk "$TEMPDIR/.merged_temp" "$TEMPDIR/$FILE" cat output "$MERGEDFILE"
            fi 
        fi
    done

    #send message back if there some attachments
    if [[ "$ATTACHMENTS" != "Did" ]]; then

        #get sender and receiver from message
        SENDER=`cat "$MESSAGEFILE" | formail -x From: | egrep -o "<.*>" | tr -d '<>'`
        RECEIVER=`cat "$MESSAGEFILE" | formail -x To: | egrep -o "<.*>" | tr -d '<>'`

        #send it back
        echo "Sending message back to $SENDER"
        curl --ssl "smtp://$SMTPSERVER/${LOGIN/@/_at_}" -u "$LOGIN:$PASS" --mail-from "$RECEIVER" --mail-rcpt "$SENDER" --upload-file "$MESSAGEFILE" -s
    fi

    #move processed message
    mv -f "$MESSAGEFILE" "$PROCESSEDDIR/$(basename "$MESSAGEFILE")"
done

#remove temp files
rm -f -r "$TEMPDIR"
