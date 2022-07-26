import glob,os, sys
import pandas as pd
import numpy as np
from PIL import Image
import pytesseract
import cv2
import shutil

def make_dir(dir_name):
	try:
		os.mkdir(dir_name)
	except FileExistsError:
		pass

os.chdir(sys.argv[1])
make_dir('anotasi')

# daftar file
# ===========
all_png     = glob.glob('*.png')

# isi data teks
# =============
rekap_hilal = pd.DataFrame()
rekap_hilal['Nama File'] = all_png
rekap_hilal['hari'] = [os.getcwd().split("\\")[-1].split('-')[-1]] * len(all_png)
rekap_hilal['bulan']= [os.getcwd().split("\\")[-1].split('-')[-2]]	* len(all_png)
rekap_hilal['tahun']= [os.getcwd().split("\\")[-1].split('-')[-3]] * len(all_png)

# baca anotasi tiap citra
# =======================
# kosongkan folder 'anotasi'
current_dir = os.getcwd()
os.chdir('anotasi')
rm = glob.glob('*')
for file in rm: os.remove(file)
os.chdir(current_dir)

print()	

sub_sec, location, flat_file, cont_streched, cont_enhanced , img_stacked = [],[],[],[],[],[]

for n_image in range(len(all_png)):	
	img = Image.open(all_png[n_image])
	img_crop = img.crop((0,0, 497, 90))
	
	current_dir = os.getcwd()
	img_crop.save('anotasi-'+all_png[n_image],quality = 100)
	pytesseract.pytesseract.tesseract_cmd = r'C:\Users\USER\AppData\Local\Tesseract-OCR\tesseract.exe'

	large_image = cv2.imread('anotasi-'+all_png[n_image])
	large_image = cv2.bitwise_not(large_image)
	large_image = cv2.resize(large_image, None, fx = 4, fy = 4)
	large_image = cv2.cvtColor(large_image, cv2.COLOR_BGR2GRAY)

	kernel = np.ones((1,1), np.uint8)

	cv2.imwrite('anotasiputih-'+all_png[n_image], large_image)
	

	'''
	# mencba parameter psm dari config
	for psm in range(10,13+1):
		config = '--oem 3 --psm %d' % psm
		txt = pytesseract.image_to_string(large_image, config = config, lang='eng')
		print('psm ', psm, ':',txt)
		whole_txt.append(txt)
	'''

	config = '--oem 3 --psm 11'
	txt = pytesseract.image_to_string(large_image, config = config, lang='eng')
	
	text = txt.split('\n')[::2]
	
	shutil.move('anotasi-'+all_png[n_image], 'anotasi')
	shutil.move('anotasiputih-'+all_png[n_image], 'anotasi')
	
	sub_sec.append(text[0].split('.')[-1].split()[0])
	location.append(text[0].split('-')[-1])
	
	if text[1] == 'No flat field correction':
		flat_file.append('No flat field correction')
	else:
		flat_file.append(text[1].split('\\')[-1])
	actions = text[2:-1]
		
	if actions[0] == 'Contrast streched': cont_streched.append('yes')
	else: cont_streched.append('no')
	
	if actions[1] == 'Contrast enhanced': cont_enhanced.append('yes')
	else: cont_enhanced.append('no')
	
	if text[-1].split()[0]== 'No':
		img_stacked.append('No image stacking')
	else:
		img_stacked.append(text[-1].split()[-1])
	
	print(str(n_image+1)+'/'+str(len(all_png)),all_png[n_image],'  ',sub_sec[-1],' ',flat_file[-1],' ',cont_streched[-1],' ',cont_enhanced[-1],' ',img_stacked[-1] )
	
jam,menit,detik = [],[],[]
for file in all_png:
	jam.append(file.split('-')[-3])
	menit.append(file.split('-')[-2])
	detik.append(file.split('-')[-1].split()[0])


detik_ss = []
for s,ss in zip(detik,sub_sec):
	detik_ss.append(s+'.'+ss)
	
	
rekap_hilal['jam']   = jam
rekap_hilal['menit'] = menit
rekap_hilal['detik'] = detik_ss


rekap_hilal['teleskop'] = ['-'] * len(all_png)
rekap_hilal['detektor'] = ['-'] * len(all_png)
rekap_hilal['filter']   = ['-'] * len(all_png)
rekap_hilal['baffle']   = ['-'] * len(all_png)

rekap_hilal['lokasi']       = location
rekap_hilal['cont_streched']= cont_streched
rekap_hilal['cont_enhanced']= cont_enhanced
rekap_hilal['file flat']    = flat_file
rekap_hilal['image stacked']= img_stacked

rekap_hilal['konjungsi']    = ['-'] * len(all_png)
rekap_hilal['usia bulan']   = ['-'] * len(all_png)
rekap_hilal['bulan Hijriah']= ['-'] * len(all_png)
rekap_hilal['keterangan']   = ['-'] * len(all_png)


# Simpan dataframe ke file excel
# ==============================
nama_file = 'Hasil Script Rekap Data Citra_'+current_dir.split("\\")[-1]+'.xlsx'
file_excel = pd.ExcelWriter(nama_file)
rekap_hilal.to_excel(file_excel)
file_excel.save()

print()
print("DataFrame berhasil ditulis ke : "+nama_file)
