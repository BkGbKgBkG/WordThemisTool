from docx2pdf import convert
import sys, fitz
from PyPDF2 import PdfFileReader
import numpy as np
from PIL import Image
import os, glob
import xlwt
from xlwt import Workbook
'''
YEU CAU CAU HINH:
1. python 3.8
2. docx2pdf
3.PyMuPDF
4.PyPDF2
5.numpy
6.Pillow
7.shutil
8.glob
9.xlwt
'''

local = os.getcwd()
url_file = os.path.join


#chuyen sang file pdf, can co COM- Ms Office 2010 hoac cao hon
def chuyenSangPDF():
	if not(os.path.exists('Temp')):
		os.mkdir('Temp')
	else:
		pass
	list_path = glob.glob('*.docx')
	Dich = url_file(local, 'Temp')
	for name in list_path:
		convert(url_file(local,name), Dich)


#Module chuyen tung trang word sang hinh anh
def chuyenSangHinhAnh(ten_file_word):	
	#chuyen file pdf sang hinh anh 
	ten_file_pdf = ten_file_word[:-4] + 'pdf'
	
	#dem so trang cua file
	pdf = PdfFileReader(open(ten_file_pdf,'rb'))
	all_pages = pdf.getNumPages()
	
	#Chuyen tu file pdf sang anh (png)
	doc = None
	try:
		doc = fitz.open(ten_file_pdf)
	except Exception as e:
		print(e)
		if doc:
			doc.close()
			exit(0)
	ten_anh = ten_file_word[:-5]+'_'
	#print (ten_anh)

	for i in range(all_pages):
		trang = doc[i]
		in4_anh = fitz.Matrix(fitz.Identity)
		in4_anh.preScale(2, 2)
		anh = trang.getPixmap(alpha = False, matrix = in4_anh)
		ten= ten_anh+str(i)+'.png'
		anh.writePNG(ten)
	return all_pages

#Module so sanh do tuong quan giua hai buc anh chup lai trang word(tham khao aivietnam.ai)
def find_corr_x_y(x,y):                                         #1
    n = len(x)                                                  #2
    prod = []
    for xi,yi in zip(x,y):                                      #3
         prod.append(xi*yi)
         
    sum_prod_x_y = sum(prod)                                    #4
    
    sum_x = sum(x)
    sum_y = sum(y)
    
    squared_sum_x = sum_x**2
    squared_sum_y = sum_y**2 
    
    x_square = []
    for xi in x:
        x_square.append(xi**2)            
    x_square_sum = sum(x_square)
    y_square=[]
    for yi in y:
        y_square.append(yi**2)        
    y_square_sum = sum(y_square)
    
    # Use formula to calculate correlation                      #5
    numerator = n*sum_prod_x_y - sum_x*sum_y
    denominator_term1 = n*x_square_sum - squared_sum_x
    denominator_term2 = n*y_square_sum - squared_sum_y
    denominator = (denominator_term1*denominator_term2)**0.5
    correlation = numerator/denominator
    
    return correlation 

def so_sanh_anh(x, y):
	#load anh va chuyen ve kieu list
	anh_1 = Image.open(x)
	anh_2 = Image.open(y)
	data_anh_1 = np.asarray(anh_1).flatten().tolist()
	data_anh_2 = np.asarray(anh_2).flatten().tolist()

	#tinh toan do tuong dong
	corr_1_2 = find_corr_x_y(data_anh_1, data_anh_2)
	return corr_1_2

#XU LI


chuyenSangPDF()
cc = True
while cc:
	file_goc = input('Nhap ten file mau: ')
	if os.path.isfile(url_file(local, file_goc)):
		cc=False


list_file_word = glob.glob('*.docx')
list_file_word.remove(file_goc)

os.chdir(url_file(local, 'Temp'))


all_pages_1 = chuyenSangHinhAnh(file_goc)
file_goc = file_goc[:-5]+'_'

print('File mau co '+str(all_pages_1) + ' trang \n')
diem_mau = [int(input('Nhap diem trang thu '+str(i+1)+': ')) for i in range(all_pages_1)]

print('\n')
def tinhKetQua(file_so_sanh):
	all_pages_2 = chuyenSangHinhAnh(file_so_sanh)
	file_so_sanh = file_so_sanh[:-5]+'_'

	ket_qua=0
	

	for i in range(min(all_pages_1, all_pages_2)):
		ten_anh_1 = file_goc +str(i)+'.png'
		ten_anh_2 = file_so_sanh +str(i)+'.png'
		try:
			corr = so_sanh_anh(ten_anh_2, ten_anh_1)
		except Exception:
			corr = 0

		os.remove(ten_anh_2)
		ket_qua += diem_mau[i]*corr
	return ket_qua

res = ''
wb = Workbook()
sheet1 = wb.add_sheet('Điểm')
col =0
row =0
sheet1.write(row, col, 'TÊN FILE')
sheet1.write(row, col+1, 'ĐIỂM')
row+=1
for name in list_file_word:
	#res += '+--------------+ \n'+ name[:-4] +' '+ str(tinhKetQua(name)) +'\n +--------------- \n \n'
	sheet1.write(row, col, name[:-5])
	sheet1.write(row, col+1, tinhKetQua(name))
	row+=1
	print(name,'___DA CHAM XONG___')

for name in glob.glob('*.png'):
	os.remove(name)


#print(os.listdir(),'\n')

os.chdir(local)
print('\nKIEM TRA FILE KetQua.xls DE BIET KET QUA')
wb.save('KetQua.xls')
