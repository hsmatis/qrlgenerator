#
#  This program creates labels for the SensTek sensor - Howard Matis
#  Version June 13, 2021
#
import pyqrcode
import xlsxwriter
#
total_codes = 5                                        # Number of lables to create
model = '{"SensTek_S":{"mn":"ST7000","sn":'    # Official Part Number
assembly_location = 100000000                          # Modify this when the assembly location is changed.
                                                       # Make sure that you do not change the number of digits
                                                       # as the we are the limit of characters for the QR code
url = pyqrcode.create('http://www.senstk.com')
# url.svg('website.svg', scale=8)
# url.eps('website.eps', scale=2)
url.png('website.png', scale=4)
#
url.svg('uca.svg', scale=4)
number = pyqrcode.create(123456789012345)
number.png('SensTek.png')
#
file_a = open('last_qr_code.txt', 'r')
line = file_a.readline()
print("The next code to use is ", line)
file_a.close()

first_id = int(line)   # Edit last_qr_code.txt to start at different number
last_id = first_id + total_codes
print("Writing",total_codes,"codes from:",first_id,"to", last_id )

# Write the xlsx file used for QR codes for DYMO printer

workbook = xlsxwriter.Workbook('dymo_input.xlsx')   # writing a csv
worksheet = workbook.add_worksheet()
row = 0
worksheet.write(row, 0, "ID")               # Write the header record
worksheet.write(row, 1, "QR-Code")

for id in range(first_id,last_id):
    row += 1
    id_str = str(assembly_location + id)
    # id_str = str(id).zfill(6)
    qrl = model + id_str + "}}"
    code = pyqrcode.create(qrl)
    worksheet.write(row, 0, id_str)
    worksheet.write(row, 1, qrl)

#    filename_1 = model + "-" + id_str + "_1.png"
#    filename_2 = model + "-" + id_str + "_2.png"
#    code.png("id/" + filename_1, scale=1)

last_qr = open('last_qr_code.txt', 'w')
print(last_id+1, file=last_qr)
workbook.close()
print("All Done")
exit()
