# A simple script to calculate BMI
from pywebio import start_server
from pywebio.input import input, FLOAT, input_group, NUMBER , file_upload , radio
from pywebio.output import put_text,put_html
import asyncio
import os
import PIL 

from btech import *

def bmi():

    try:
        shutil.rmtree('media')
    except:
        pass

    os.mkdir('media')
    
    try:
        shutil.rmtree('transcriptsIITP')
    except:
        pass

    os.mkdir('transcriptsIITP')
    
    put_html('<h2>Transcript Generator</h2>')

    data = radio("Choose your preference :", options=['Generate all Transcripts', 'Enter a range of Roll Numbers'])
    
    if(data == 'Generate all Transcripts'):
        info = input_group("Enter Data",[
        file_upload("Select a Seal:", accept="image/*",name="Seal"),
        file_upload("Select a Signature:", accept="image/*",name="Signature")
        ])

        try:
            open('media/seal.png', 'wb').write(info['Seal']['content']) 
        except:
            pass
        try:
            open('media/sign.png', 'wb').write(info['Signature']['content']) 
        except:
            pass

        put_text('All Generated')

        btech_transcript('*', '*')
        print('All Generated')
        
        

    if(data == 'Enter a range of Roll Numbers'):
        info = input_group("Enter Data",[
        file_upload("Select a Seal:", accept="image/*",name="Seal"),
        file_upload("Select a Signature:", accept="image/*",name="Signature"),
        input('From Roll No', name='roll_1'),
        input('To Roll No', name='roll_2')
        ])

        try:
            open('media/seal.png', 'wb').write(info['Seal']['content']) 
        except:
            pass
        try:
            open('media/sign.png', 'wb').write(info['Signature']['content']) 
        except:
            pass

        btech_transcript(info['roll_1'].upper(), info['roll_2'].upper())
        list = []
        for i in range(int(info['roll_1'][6:]), int(info['roll_2'][6:]) + 1):
            rollfile = 'transcriptsIITP/' + info['roll_1'].upper()[:6] + str(i).zfill(2) + '.pdf'
            if not os.path.exists(rollfile):
                list.append(info['roll_1'].upper()[:6] + str(i).zfill(2))

        put_text('From Roll: %s to Roll %s is generated' % (info['roll_1'].upper(), info['roll_2'].upper()))
        if len(list) > 0:
            put_text('''These Roll nos. don't exist :''')
            put_text(list)
    
    

if __name__=='__main__':
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    start_server(bmi, port=8080, debug=True)