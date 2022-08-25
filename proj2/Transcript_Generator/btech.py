from fpdf import FPDF
import os
import shutil
import csv

from datetime import datetime
import pytz
IST = pytz.timezone('Asia/Kolkata')

def btech_transcript(start, end):

    try:
        shutil.rmtree('transcriptsIITP')
    except:
        pass

    os.mkdir('transcriptsIITP')

    grade_eq = {'AA':10,'AB':9,'BB':8,' BB':8,'BC':7,'CC':6,'CD':5,'DD':4,'DD*':4,'F':0,'F*':0,'I':0,'I*':0}

    course = {'CS':'Computer Science and Engineering',
              'EE':'Electrical Engineering',
              'ME':'Mechanical Engineering',
              'CB':'Chemical and Biochemical Engineering',
              'CE':'Civil and Environmental Engineering',
              'MM':'Metallurgical and Materials Engineering'}

    # Maping grades...
    with open('sample_input/grades.csv', 'r') as csvfile:

        reader = csv.DictReader(csvfile)

        grade_record = {}
        for row in reader:
            
            list = []
            list.append(row['Roll'])
            list.append(row['Sem'])
            list.append(row['SubCode'])
            list.append(row['Credit'])
            list.append(row['Grade'])
            list.append(row['Sub_Type'])

            if(grade_record.get(row['Roll']) == None):
                grade_record[row['Roll']] = [list]
            else:
                grade_record[row['Roll']].append(list)

    # Maping names-roll...
    with open('sample_input/names-roll.csv', 'r') as csvfile:

        reader = csv.DictReader(csvfile)

        name = {}
        for row in reader:
            name[row['Roll']] = row['Name']

    # Maping subjects_master...
    with open('sample_input/subjects_master.csv', 'r') as csvfile:

        csvreader = csv.DictReader(csvfile)

        subject = {}
        for row in csvreader:
            
            list = []
            list.append(row['subname'])
            list.append(row['ltp'])

            subject[row['subno']] = list


    # Storing result of each student in a nested dictionary...

    result = {}

    for roll in grade_record:
        
        if (start == '*') or (start <= roll <= end):
        
            stud_result = {}

            for row in grade_record[roll]:

                sem = row[1]
                subno = row[2]
                credit = row[3]
                grade = row[4]
                sub_type = row[5]

                list = []
                list.append(subno)
                list.append(subject[subno][0])
                list.append(subject[subno][1])
                list.append(credit)
                # list.append(sub_type)
                list.append(grade)

                if(stud_result.get(sem) == None):
                    stud_result[sem] = [list]
                else:
                    stud_result[sem].append(list)
            
            result[roll] = stud_result


    # df = pd.DataFrame(result[roll]['1'])


    # print(result[roll]['1'])
    #print(result[roll])


    res_dict = {}

    for roll in grade_record:

        if (start == '*') or (start <= roll <= end):
            
            list = []
            # Calculations for Overall sheet...
            semester = []
            spi = []
            cpi = []
            cred_taken = []
            total_cred_taken = []
            total_cred_sum = 0
            cpi_sum = 0

            for i in result[roll]:

                semester.append(i)
                spi_sum = 0
                cred_sum = 0

                for row in result[roll][i]:
                    marks = float(grade_eq[row[4]])
                    cred = float(row[3])

                    spi_sum += marks*cred
                    cred_sum += cred
                
                total_cred_sum += cred_sum
                cpi_sum += (spi_sum/cred_sum)*cred_sum
                
                spi.append(round(spi_sum/cred_sum, 2))
                cpi.append(round(cpi_sum/total_cred_sum, 2))
                cred_taken.append(cred_sum)
                total_cred_taken.append(total_cred_sum)
            
            for i in range(len(spi)):
                list.append([int(cred_taken[i]),spi[i],cpi[i]])
                
            res_dict[roll] = list

    # print(res_dict[roll])


    for roll in result:

        if roll[2:4] == '01':

            data = result[roll]

            header = ["Sub. Code", "Subject Name", "L-T-P", "CRD", "GRD"]

            pdf = FPDF()
            pdf.add_page(format='A3')
            pdf.set_font("Arial","",6)

            x0 = pdf.x
            y0 = pdf.y

            pdf.image('iitplogo.png', w=277, h=30)
            pdf.x = x0
            pdf.y = y0
            pdf.cell(277, 30, '', 1, 1, 'C')
            pdf.y = 20
            y1 = 50
            ymax = y1
            line_height = pdf.font_size * 2

            #######################################

            pdf.set_font("Arial","B",9)
            pdf.y = 44
            pdf.x = x0 + 52
            pdf.cell(15, line_height, "Roll No:")
            pdf.x = x0 + 112
            pdf.cell(15, line_height, "Name:")
            pdf.x = x0 + 180
            pdf.cell(15, line_height, "Year of Admission:")
            pdf.y = 51
            pdf.x = x0 + 52
            pdf.cell(15, line_height, "Programme:")
            pdf.x = x0 + 112
            pdf.cell(15, line_height, "Course:")

            pdf.set_font("Arial","",9)
            pdf.y = 44
            pdf.x = x0 + 72
            pdf.cell(15, line_height, roll)
            pdf.x = x0 + 125
            pdf.cell(15, line_height, name[roll])
            pdf.x = x0 + 211
            pdf.cell(15, line_height, "20" + roll[:2])
            pdf.y = 51
            pdf.x = x0 + 72
            pdf.cell(15, line_height, "Bachelor of Technology")
            pdf.x = x0 + 125
            pdf.cell(15, line_height, course[roll[4:6]])
            pdf.set_font("Arial","",6)
            #######################################

            pdf.y = 42
            pdf.x = x0 + 50
            pdf.cell(177, 15, '', 1, 1, 'C')
            pdf.ln(line_height)

            pdf.set_font("Arial","B",6)

            if '1' in data.keys():
                pdf.y = y1 + 12
                pdf.x += 4
                pdf.set_font("Arial","BU",9)
                pdf.cell(15, line_height, "Semester 1")
                pdf.set_font("Arial","B",6)
                pdf.ln(line_height)

                pdf.y = y1 + 20
                data['1'] = [header] + data['1']
                for row in data['1']:
                    pdf.x += 4
                    pdf.multi_cell(13, line_height, row[0], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(50, line_height, row[1], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(9, line_height, row[2], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(7, line_height, row[3], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","B",6)
                    pdf.multi_cell(7, line_height, row[4], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","",6)
                    pdf.ln(line_height)
                pdf.x += 4
                pdf.y += 2
                pdf.set_font("Arial","B",6)
                text = "Credits Taken:  %d   Credits Cleared:  %d   SPI:  %.2f   CPI:  %.2f"%(res_dict[roll][0][0], res_dict[roll][0][0], res_dict[roll][0][1], res_dict[roll][0][2])
                pdf.multi_cell(70, pdf.font_size * 2.5, text, border=1, align='C', ln=3, max_line_height=pdf.font_size)
                pdf.ln(line_height)

                ymax = max(ymax, pdf.y)

            if '2' in data.keys():
                pdf.y = y1 + 12
                pdf.x += 95
                pdf.set_font("Arial","BU",9)
                pdf.cell(15, line_height, "Semester 2")
                pdf.set_font("Arial","B",6)
                pdf.ln(line_height)

                pdf.y = y1 + 20
                data['2'] = [header] + data['2']
                for row in data['2']:
                    pdf.x += 95
                    pdf.multi_cell(13, line_height, row[0], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(50, line_height, row[1], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(9, line_height, row[2], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(7, line_height, row[3], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","B",6)
                    pdf.multi_cell(7, line_height, row[4], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","",6)
                    pdf.ln(line_height)
                pdf.x += 95
                pdf.y += 2
                pdf.set_font("Arial","B",6)
                text = "Credits Taken:  %d   Credits Cleared:  %d   SPI:  %.2f   CPI:  %.2f"%(res_dict[roll][1][0], res_dict[roll][1][0], res_dict[roll][1][1], res_dict[roll][1][2])
                pdf.multi_cell(70, pdf.font_size * 2.5, text, border=1, align='C', ln=3, max_line_height=pdf.font_size)
                pdf.ln(line_height)

                ymax = max(ymax, pdf.y)

            if '3' in data.keys():
                pdf.y = y1 + 12
                pdf.x += 186
                pdf.set_font("Arial","BU",9)
                pdf.cell(15, line_height, "Semester 3")
                pdf.set_font("Arial","B",6)
                pdf.ln(line_height)

                pdf.y = y1 + 20
                data['3'] = [header] + data['3']
                for row in data['3']:
                    pdf.x += 186
                    pdf.multi_cell(13, line_height, row[0], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(50, line_height, row[1], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(9, line_height, row[2], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(7, line_height, row[3], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","B",6)
                    pdf.multi_cell(7, line_height, row[4], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","",6)
                    pdf.ln(line_height)
                pdf.x += 186
                pdf.y += 2
                pdf.set_font("Arial","B",6)
                text = "Credits Taken:  %d   Credits Cleared:  %d   SPI:  %.2f   CPI:  %.2f"%(res_dict[roll][2][0], res_dict[roll][2][0], res_dict[roll][2][1], res_dict[roll][2][2])
                pdf.multi_cell(70, pdf.font_size * 2.5, text, border=1, align='C', ln=3, max_line_height=pdf.font_size)
                pdf.ln(line_height)

                ymax = max(ymax, pdf.y)
            
            if '1' in data.keys():
                y2 = ymax
                pdf.y = y2 + 8
                pdf.cell(277, 0, '', 1, 1, 'C')
            ######################################################

            if '4' in data.keys():
                pdf.y = y2 + 12
                pdf.x += 4
                pdf.set_font("Arial","BU",9)
                pdf.cell(15, line_height, "Semester 4")
                pdf.set_font("Arial","B",6)
                pdf.ln(line_height)

                pdf.y = y2 + 20
                data['4'] = [header] + data['4']
                for row in data['4']:
                    pdf.x += 4
                    pdf.multi_cell(13, line_height, row[0], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(50, line_height, row[1], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(9, line_height, row[2], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(7, line_height, row[3], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","B",6)
                    pdf.multi_cell(7, line_height, row[4], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","",6)
                    pdf.ln(line_height)
                pdf.x += 4
                pdf.y += 2
                pdf.set_font("Arial","B",6)
                text = "Credits Taken:  %d   Credits Cleared:  %d   SPI:  %.2f   CPI:  %.2f"%(res_dict[roll][3][0], res_dict[roll][3][0], res_dict[roll][3][1], res_dict[roll][3][2])
                pdf.multi_cell(70, pdf.font_size * 2.5, text, border=1, align='C', ln=3, max_line_height=pdf.font_size)
                pdf.ln(line_height)

                ymax = max(ymax, pdf.y)

            if '5' in data.keys():
                pdf.y = y2 + 12
                pdf.x += 95
                pdf.set_font("Arial","BU",9)
                pdf.cell(15, line_height, "Semester 5")
                pdf.set_font("Arial","B",6)
                pdf.ln(line_height)

                pdf.y = y2 + 20
                data['5'] = [header] + data['5']
                for row in data['5']:
                    pdf.x += 95
                    pdf.multi_cell(13, line_height, row[0], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(50, line_height, row[1], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(9, line_height, row[2], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(7, line_height, row[3], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","B",6)
                    pdf.multi_cell(7, line_height, row[4], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","",6)
                    pdf.ln(line_height)
                pdf.x += 95
                pdf.y += 2
                pdf.set_font("Arial","B",6)
                text = "Credits Taken:  %d   Credits Cleared:  %d   SPI:  %.2f   CPI:  %.2f"%(res_dict[roll][4][0], res_dict[roll][4][0], res_dict[roll][4][1], res_dict[roll][4][2])
                pdf.multi_cell(70, pdf.font_size * 2.5, text, border=1, align='C', ln=3, max_line_height=pdf.font_size)
                pdf.ln(line_height)

                ymax = max(ymax, pdf.y)

            if '6' in data.keys():
                pdf.y = y2 + 12
                pdf.x += 186
                pdf.set_font("Arial","BU",9)
                pdf.cell(15, line_height, "Semester 6")
                pdf.set_font("Arial","B",6)
                pdf.ln(line_height)

                pdf.y = y2 + 20
                data['6'] = [header] + data['6']
                for row in data['6']:
                    pdf.x += 186
                    pdf.multi_cell(13, line_height, row[0], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(50, line_height, row[1], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(9, line_height, row[2], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(7, line_height, row[3], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","B",6)
                    pdf.multi_cell(7, line_height, row[4], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","",6)
                    pdf.ln(line_height)
                pdf.x += 186
                pdf.y += 2
                pdf.set_font("Arial","B",6)
                text = "Credits Taken:  %d   Credits Cleared:  %d   SPI:  %.2f   CPI:  %.2f"%(res_dict[roll][5][0], res_dict[roll][5][0], res_dict[roll][5][1], res_dict[roll][5][2])
                pdf.multi_cell(70, pdf.font_size * 2.5, text, border=1, align='C', ln=3, max_line_height=pdf.font_size)
                pdf.ln(line_height)

                ymax = max(ymax, pdf.y)
            
            if '4' in data.keys():
                y3 = ymax
                pdf.y = y3 + 8
                pdf.cell(277, 0, '', 1, 1, 'C')
            ######################################################

            if '7' in data.keys():
                pdf.y = y3 + 12
                pdf.x += 4
                pdf.set_font("Arial","BU",9)
                pdf.cell(15, line_height, "Semester 7")
                pdf.set_font("Arial","B",6)
                pdf.ln(line_height)

                pdf.y = y3 + 20
                data['7'] = [header] + data['7']
                for row in data['7']:
                    pdf.x += 4
                    pdf.multi_cell(13, line_height, row[0], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(50, line_height, row[1], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(9, line_height, row[2], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(7, line_height, row[3], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","B",6)
                    pdf.multi_cell(7, line_height, row[4], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","",6)
                    pdf.ln(line_height)
                pdf.x += 4
                pdf.y += 2
                pdf.set_font("Arial","B",6)
                text = "Credits Taken:  %d   Credits Cleared:  %d   SPI:  %.2f   CPI:  %.2f"%(res_dict[roll][6][0], res_dict[roll][6][0], res_dict[roll][6][1], res_dict[roll][6][2])
                pdf.multi_cell(70, pdf.font_size * 2.5, text, border=1, align='C', ln=3, max_line_height=pdf.font_size)
                pdf.ln(line_height)

                ymax = max(ymax, pdf.y)

            if '8' in data.keys():
                pdf.y = y3 + 12
                pdf.x += 95
                pdf.set_font("Arial","BU",9)
                pdf.cell(15, line_height, "Semester 8")
                pdf.set_font("Arial","B",6)
                pdf.ln(line_height)

                pdf.y = y3 + 20
                data['8'] = [header] + data['8']
                for row in data['8']:
                    pdf.x += 95
                    pdf.multi_cell(13, line_height, row[0], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(50, line_height, row[1], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(9, line_height, row[2], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(7, line_height, row[3], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","B",6)
                    pdf.multi_cell(7, line_height, row[4], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","",6)
                    pdf.ln(line_height)
                pdf.x += 95
                pdf.y += 2
                pdf.set_font("Arial","B",6)
                text = "Credits Taken:  %d   Credits Cleared:  %d   SPI:  %.2f   CPI:  %.2f"%(res_dict[roll][7][0], res_dict[roll][7][0], res_dict[roll][7][1], res_dict[roll][7][2])
                pdf.multi_cell(70, pdf.font_size * 2.5, text, border=1, align='C', ln=3, max_line_height=pdf.font_size)
                pdf.ln(line_height)

                ymax = max(ymax, pdf.y)

            if '7' in data.keys():
                pdf.y = ymax + 10
                pdf.cell(277, 0, '', 1, 1, 'C')

            now = datetime.now(IST)
            date_time = str(now.day).zfill(2) + ' ' + now.strftime("%B")[:3] + ' ' + str(now.year) + ', ' + str(now.hour).zfill(2) + ":" + str(now.minute).zfill(2)

            pdf.set_font("Arial","",9)
            pdf.y = ymax + 40
            pdf.x = x0 + 29
            pdf.cell(15, line_height, date_time)
            pdf.set_font("Arial","B",9)
            pdf.x = x0 + 3
            pdf.cell(15, line_height, "Date Generated:")
            pdf.y = ymax + 55
            pdf.x = x0 + 220
            pdf.cell(15, line_height, "Assistant Registrar (Academic)")
            
            try:
                pdf.y = ymax + 25
                pdf.x = x0 + 225
                pdf.image('media/sign.png', w=35, h=25)
            except:
                pass
            try:
                pdf.y = ymax + 25
                pdf.x = x0 + 110
                pdf.image('media/seal.png', w=30, h=30)
            except:
                pass

            pdf.x = x0
            pdf.y = y0
            pdf.cell(277, ymax + 70, '', 1, 1, 'C')

            # for row in data:
            #     pdf.x = 50
            #     for datum in row:
            #         pdf.multi_cell(col_width, line_height, datum, border=1, ln=3, max_line_height=pdf.font_size)
            #     pdf.ln(line_height)

            pdf.output('transcriptsIITP/' + roll + '.pdf')


        else:

            data = result[roll]

            header = ["Sub. Code", "Subject Name", "L-T-P", "CRD", "GRD"]

            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial","",6)


            x0 = pdf.x
            y0 = pdf.y

            pdf.image('iitplogo.png', w=190, h=30)
            pdf.x = x0
            pdf.y = y0
            pdf.cell(190, 30, '', 1, 1, 'C')
            pdf.y = 20
            y1 = 50
            ymax = y1
            line_height = pdf.font_size * 2

            #######################################

            pdf.set_font("Arial","B",9)
            pdf.y = 44
            pdf.x = x0 + 9
            pdf.cell(15, line_height, "Roll No:")
            pdf.x = x0 + 69
            pdf.cell(15, line_height, "Name:")
            pdf.x = x0 + 137
            pdf.cell(15, line_height, "Year of Admission:")
            pdf.y = 51
            pdf.x = x0 + 9
            pdf.cell(15, line_height, "Programme:")
            pdf.x = x0 + 69
            pdf.cell(15, line_height, "Course:")

            pdf.set_font("Arial","",9)
            pdf.y = 44
            pdf.x = x0 + 29
            pdf.cell(15, line_height, roll)
            pdf.x = x0 + 82
            pdf.cell(15, line_height, name[roll])
            pdf.x = x0 + 168
            pdf.cell(15, line_height, "20" + roll[:2])
            pdf.y = 51
            pdf.x = x0 + 29
            pdf.cell(15, line_height, "Bachelor of Technology")
            pdf.x = x0 + 82
            pdf.cell(15, line_height, roll[4:6])
            pdf.set_font("Arial","",6)
            #######################################

            pdf.y = 42
            pdf.x = x0 + 7
            pdf.cell(176, 15, '', 1, 1, 'C')
            pdf.ln(line_height)

            pdf.set_font("Arial","B",6)

            if '1' in data.keys():
                pdf.y = y1 + 12
                pdf.x += 4
                pdf.set_font("Arial","BU",9)
                pdf.cell(15, line_height, "Semester 1")
                pdf.set_font("Arial","B",6)
                pdf.ln(line_height)

                pdf.y = y1 + 20
                data['1'] = [header] + data['1']
                for row in data['1']:
                    pdf.x += 4
                    pdf.multi_cell(13, line_height, row[0], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(50, line_height, row[1], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(9, line_height, row[2], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(7, line_height, row[3], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","B",6)
                    pdf.multi_cell(7, line_height, row[4], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","",6)
                    pdf.ln(line_height)
                pdf.x += 4
                pdf.y += 2
                pdf.set_font("Arial","B",6)
                text = "Credits Taken:  %d   Credits Cleared:  %d   SPI:  %.2f   CPI:  %.2f"%(res_dict[roll][0][0], res_dict[roll][0][0], res_dict[roll][0][1], res_dict[roll][0][2])
                pdf.multi_cell(70, pdf.font_size * 2.5, text, border=1, align='C', ln=3, max_line_height=pdf.font_size)
                pdf.ln(line_height)

                ymax = max(ymax, pdf.y)

            if '2' in data.keys():
                pdf.y = y1 + 12
                pdf.x += 95
                pdf.set_font("Arial","BU",9)
                pdf.cell(15, line_height, "Semester 2")
                pdf.set_font("Arial","B",6)
                pdf.ln(line_height)

                pdf.y = y1 + 20
                data['2'] = [header] + data['2']
                for row in data['2']:
                    pdf.x += 95
                    pdf.multi_cell(13, line_height, row[0], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(50, line_height, row[1], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(9, line_height, row[2], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(7, line_height, row[3], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","B",6)
                    pdf.multi_cell(7, line_height, row[4], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","",6)
                    pdf.ln(line_height)
                pdf.x += 95
                pdf.y += 2
                pdf.set_font("Arial","B",6)
                text = "Credits Taken:  %d   Credits Cleared:  %d   SPI:  %.2f   CPI:  %.2f"%(res_dict[roll][1][0], res_dict[roll][1][0], res_dict[roll][1][1], res_dict[roll][1][2])
                pdf.multi_cell(70, pdf.font_size * 2.5, text, border=1, align='C', ln=3, max_line_height=pdf.font_size)
                pdf.ln(line_height)

                ymax = max(ymax, pdf.y)

            # if '3' in data.keys():
            #     pdf.y = y1 + 12
            #     pdf.x += 186
            #     pdf.set_font("Arial","BU",9)
            #     pdf.cell(15, line_height, "Semester 3")
            #     pdf.set_font("Arial","B",6)
            #     pdf.ln(line_height)

            #     pdf.y = y1 + 20
            #     data['3'] = [header] + data['3']
            #     for row in data['3']:
            #         pdf.x += 186
            #         pdf.multi_cell(13, line_height, row[0], border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #         pdf.multi_cell(50, line_height, row[1], border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #         pdf.multi_cell(9, line_height, row[2], border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #         pdf.multi_cell(7, line_height, row[3], border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #         pdf.set_font("Arial","B",6)
            #         pdf.multi_cell(7, line_height, row[4], border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #         pdf.set_font("Arial","",6)
            #         pdf.ln(line_height)
            #     pdf.x += 186
            #     pdf.y += 2
            #     pdf.set_font("Arial","B",6)
            #     text = "Credits Taken:  %d   Credits Cleared:  %d   SPI:  %.2f   CPI:  %.2f"%(res_dict[roll][2][0], res_dict[roll][2][0], res_dict[roll][2][1], res_dict[roll][2][2])
            #     pdf.multi_cell(70, pdf.font_size * 2.5, text, border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #     pdf.ln(line_height)

            #     ymax = max(ymax, pdf.y)
            
            if '1' in data.keys():
                y2 = ymax
                pdf.y = y2 + 8
                pdf.cell(190, 0, '', 1, 1, 'C')
            ######################################################

            if '3' in data.keys():
                pdf.y = y2 + 12
                pdf.x += 4
                pdf.set_font("Arial","BU",9)
                pdf.cell(15, line_height, "Semester 3")
                pdf.set_font("Arial","B",6)
                pdf.ln(line_height)

                pdf.y = y2 + 20
                data['3'] = [header] + data['3']
                for row in data['3']:
                    pdf.x += 4
                    pdf.multi_cell(13, line_height, row[0], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(50, line_height, row[1], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(9, line_height, row[2], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(7, line_height, row[3], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","B",6)
                    pdf.multi_cell(7, line_height, row[4], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","",6)
                    pdf.ln(line_height)
                pdf.x += 4
                pdf.y += 2
                pdf.set_font("Arial","B",6)
                text = "Credits Taken:  %d   Credits Cleared:  %d   SPI:  %.2f   CPI:  %.2f"%(res_dict[roll][3][0], res_dict[roll][3][0], res_dict[roll][3][1], res_dict[roll][3][2])
                pdf.multi_cell(70, pdf.font_size * 2.5, text, border=1, align='C', ln=3, max_line_height=pdf.font_size)
                pdf.ln(line_height)

                ymax = max(ymax, pdf.y)

            if '4' in data.keys():
                pdf.y = y2 + 12
                pdf.x += 95
                pdf.set_font("Arial","BU",9)
                pdf.cell(15, line_height, "Semester 4")
                pdf.set_font("Arial","B",6)
                pdf.ln(line_height)

                pdf.y = y2 + 20
                data['4'] = [header] + data['4']
                for row in data['4']:
                    pdf.x += 95
                    pdf.multi_cell(13, line_height, row[0], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(50, line_height, row[1], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(9, line_height, row[2], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.multi_cell(7, line_height, row[3], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","B",6)
                    pdf.multi_cell(7, line_height, row[4], border=1, align='C', ln=3, max_line_height=pdf.font_size)
                    pdf.set_font("Arial","",6)
                    pdf.ln(line_height)
                pdf.x += 95
                pdf.y += 2
                pdf.set_font("Arial","B",6)
                text = "Credits Taken:  %d   Credits Cleared:  %d   SPI:  %.2f   CPI:  %.2f"%(res_dict[roll][4][0], res_dict[roll][4][0], res_dict[roll][4][1], res_dict[roll][4][2])
                pdf.multi_cell(70, pdf.font_size * 2.5, text, border=1, align='C', ln=3, max_line_height=pdf.font_size)
                pdf.ln(line_height)

                ymax = max(ymax, pdf.y)

            # if '6' in data.keys():
            #     pdf.y = y2 + 12
            #     pdf.x += 186
            #     pdf.set_font("Arial","BU",9)
            #     pdf.cell(15, line_height, "Semester 6")
            #     pdf.set_font("Arial","B",6)
            #     pdf.ln(line_height)

            #     pdf.y = y2 + 20
            #     data['6'] = [header] + data['6']
            #     for row in data['6']:
            #         pdf.x += 186
            #         pdf.multi_cell(13, line_height, row[0], border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #         pdf.multi_cell(50, line_height, row[1], border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #         pdf.multi_cell(9, line_height, row[2], border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #         pdf.multi_cell(7, line_height, row[3], border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #         pdf.set_font("Arial","B",6)
            #         pdf.multi_cell(7, line_height, row[4], border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #         pdf.set_font("Arial","",6)
            #         pdf.ln(line_height)
            #     pdf.x += 186
            #     pdf.y += 2
            #     pdf.set_font("Arial","B",6)
            #     text = "Credits Taken:  %d   Credits Cleared:  %d   SPI:  %.2f   CPI:  %.2f"%(res_dict[roll][5][0], res_dict[roll][5][0], res_dict[roll][5][1], res_dict[roll][5][2])
            #     pdf.multi_cell(70, pdf.font_size * 2.5, text, border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #     pdf.ln(line_height)

            #     ymax = max(ymax, pdf.y)
            
            if '3' in data.keys():
                y3 = ymax
                pdf.y = y3 + 8
                pdf.cell(190, 0, '', 1, 1, 'C')
            ######################################################

            # if '7' in data.keys():
            #     pdf.y = y3 + 12
            #     pdf.x += 4
            #     pdf.set_font("Arial","BU",9)
            #     pdf.cell(15, line_height, "Semester 7")
            #     pdf.set_font("Arial","B",6)
            #     pdf.ln(line_height)

            #     pdf.y = y3 + 20
            #     data['7'] = [header] + data['7']
            #     for row in data['7']:
            #         pdf.x += 4
            #         pdf.multi_cell(13, line_height, row[0], border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #         pdf.multi_cell(50, line_height, row[1], border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #         pdf.multi_cell(9, line_height, row[2], border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #         pdf.multi_cell(7, line_height, row[3], border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #         pdf.set_font("Arial","B",6)
            #         pdf.multi_cell(7, line_height, row[4], border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #         pdf.set_font("Arial","",6)
            #         pdf.ln(line_height)
            #     pdf.x += 4
            #     pdf.y += 2
            #     pdf.set_font("Arial","B",6)
            #     text = "Credits Taken:  %d   Credits Cleared:  %d   SPI:  %.2f   CPI:  %.2f"%(res_dict[roll][6][0], res_dict[roll][6][0], res_dict[roll][6][1], res_dict[roll][6][2])
            #     pdf.multi_cell(70, pdf.font_size * 2.5, text, border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #     pdf.ln(line_height)

            #     ymax = max(ymax, pdf.y)

            # if '8' in data.keys():
            #     pdf.y = y3 + 12
            #     pdf.x += 95
            #     pdf.set_font("Arial","BU",9)
            #     pdf.cell(15, line_height, "Semester 8")
            #     pdf.set_font("Arial","B",6)
            #     pdf.ln(line_height)

            #     pdf.y = y3 + 20
            #     data['8'] = [header] + data['8']
            #     for row in data['8']:
            #         pdf.x += 95
            #         pdf.multi_cell(13, line_height, row[0], border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #         pdf.multi_cell(50, line_height, row[1], border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #         pdf.multi_cell(9, line_height, row[2], border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #         pdf.multi_cell(7, line_height, row[3], border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #         pdf.set_font("Arial","B",6)
            #         pdf.multi_cell(7, line_height, row[4], border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #         pdf.set_font("Arial","",6)
            #         pdf.ln(line_height)
            #     pdf.x += 95
            #     pdf.y += 2
            #     pdf.set_font("Arial","B",6)
            #     text = "Credits Taken:  %d   Credits Cleared:  %d   SPI:  %.2f   CPI:  %.2f"%(res_dict[roll][7][0], res_dict[roll][7][0], res_dict[roll][7][1], res_dict[roll][7][2])
            #     pdf.multi_cell(70, pdf.font_size * 2.5, text, border=1, align='C', ln=3, max_line_height=pdf.font_size)
            #     pdf.ln(line_height)

            #     ymax = max(ymax, pdf.y)

            # if '7' in data.keys():
            #     pdf.y = ymax + 10
            #     pdf.cell(277, 0, '', 1, 1, 'C')

            now = datetime.now(IST)
            date_time = str(now.day).zfill(2) + ' ' + now.strftime("%B")[:3] + ' ' + str(now.year) + ', ' + str(now.hour).zfill(2) + ":" + str(now.minute).zfill(2)

            pdf.set_font("Arial","",9)
            pdf.y = ymax + 40
            pdf.x = x0 + 27
            pdf.cell(15, line_height, date_time)
            pdf.set_font("Arial","B",9)
            pdf.x = x0 + 3
            pdf.cell(15, line_height, "Date Generated:")
            pdf.y = ymax + 55
            pdf.x = x0 + 138
            pdf.cell(15, line_height, "Assistant Registrar (Academic)")

            try:
                pdf.y = ymax + 25
                pdf.x = x0 + 143
                pdf.image('media/sign.png', w=35, h=25)
            except:
                pass
            try:
                pdf.y = ymax + 25
                pdf.x = x0 + 80
                pdf.image('media/seal.png', w=30, h=30)
            except:
                pass

            pdf.x = x0
            pdf.y = y0
            pdf.cell(190, ymax + 70, '', 1, 1, 'C')

            # for row in data:
            #     pdf.x = 50
            #     for datum in row:
            #         pdf.multi_cell(col_width, line_height, datum, border=1, ln=3, max_line_height=pdf.font_size)
            #     pdf.ln(line_height)

            pdf.output('transcriptsIITP/' + roll + '.pdf')
