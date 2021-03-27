# How to use webhook with Python to send messages to Google ChatRoom:
# https://developers.google.com/hangouts/chat/quickstart/incoming-bot-python
import datetime
import time
from datetime import date
from json import dumps
from httplib2 import Http
from conf import conf

def SendChatMessage(sMessage, iRoomOrder):
    url = conf.arrWebhookURL[iRoomOrder]
    sMessageToSend = sMessage
    bot_message = {
        'text' : sMessageToSend}

    message_headers = {'Content-Type': 'application/json; charset=UTF-8'}

    http_obj = Http()

    response = http_obj.request(
        uri=url,
        method='POST',
        headers=message_headers,
        body=dumps(bot_message),
    )

    # print("Message {} has been sent to Room {}".format(sMessage, iRoomOrder))



def main() :
    for i in range(1,1000,1):
        sText = datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")
        if (i%2) == 0:
            SendChatMessage(sText, 7)
        else:
            SendChatMessage("*"+sText+"*", 7)
        time.sleep(30)

if __name__ == '__main__':
    main()
