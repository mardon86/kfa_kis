import re, string, os, time, math
from operator import itemgetter
from decimal import Decimal

def load_raw_master(raw_master_file):
  with open(raw_master_file,'r') as rmf:
    a = string.split(re.sub('\r','', rmf.read()),'\n')
    return a
  

def get_pattern(loaded_raw_master_file):
  amount_of_dashes = []
  for line in loaded_raw_master_file:
    if '----' in line:
      for i in string.split(line,' '):
        amount_of_dashes.append(len(i))
      break
  return amount_of_dashes


def create_master_list(loaded_raw_master_file, pattern):
  master_list = []
  for line in loaded_raw_master_file:
    if len(line) > 0:
      if line[0].isdigit():
        line_list = []
        curs = 0
        for col in pattern:
          line_list.append(re.sub('^ *| *$','',line[curs:curs+col]))
          curs = curs + col + 1
        master_list.append(line_list)
  return master_list


def get_big_80(the_list):
  tl = the_list
  sum_values = 0
  for i,j,k in tl:
    sum_values += k
  big_80 = []
  eighty_percent_of_sum_values = Decimal('0.8') * sum_values
  sum_big80 = 0
  for i in tl:
    if sum_big80 < eighty_percent_of_sum_values:
      big_80.append(i)
      sum_big80 += i[2]
    else:
      break
  return big_80

    
if __name__ == '__main__':

  temp_dir = 'tempdir'
#  temp_dir = '.temp_'+''.join(random.choice('ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789') for _ in range(8))
#  os.mkdir(temp_dir)
  
#  os.system('db2 connect to app2013')
#  os.system('db2 -z{0}TBMASOBAT.TXT select KODE_OBAT,NAMA_OBAT,LOKASI,SATUAN_OBAT,QTYAKHIR_OBAT,HNAPPNREG_OBAT,FAKTOR3,HJAH,DISCH,TANGGAL1,TANGGAL2 from TBMASOBAT'.format(temp_dir + os.path.sep))
  
#  last_rec = os.popen('db2 select BLNTHN_RESEP from TBHISDETILJUAL order by TANGGAL_RESEP desc fetch first 1 row only').read()[30:34]
#  last_month = last_rec[0:2]
#  last_year = last_rec[2:]

#  if last_month == "02":
#    first_month = "12"
#    first_year = "{0}".format(int(last_year) - 1)
#    sec_month = "01"
#    sec_year = last_yearlist_tbhisdetiljual1
#  elif last_month == "01":
#    first_month = "11"
#    first_year = "{0}".format(int(last_year) - 1)
#    sec_month = "12"
#    sec_year = "{0}".format(int(last_year) - 1)
#  elif last_month == "12":
#    first_month = "10"
#    first_year = last_year
#    sec_month = "11"
#    sec year = last_year
#  elif last_month == "11":
#    first_month = "09"
#    first_year = last_year
#    sec_month = "10"
#    sec_year = last_year
#  else:
#    first_month = "0{0}".format(int(last_month) - 2)
#    first_year = last_year
#    sec_month = "0{0}".format(int(last_month) - 1)
#    sec_year = last_year

#  os.system('db2 -z{0}TBHISDETILJUAL_1.TXT select KODE_OBAT,NAMA_OBAT,QTYJUAL_OBAT,JMLHRG_NETTO from TBHISDETILJUAL where BLNTHN_RESEP = \'{1}{2}\''.format(temp_dir + os.path.sep, first_month, first_year))
#  os.system('db2 -z{0}TBHISDETILJUAL_2.TXT select KODE_OBAT,NAMA_OBAT,QTYJUAL_OBAT,JMLHRG_NETTO from TBHISDETILJUAL where BLNTHN_RESEP = \'{1}{2}\''.format(temp_dir + os.path.sep, sec_month, sec_year))
#  os.system('db2 -z{0}TBHISDETILJUAL_3.TXT select KODE_OBAT,NAMA_OBAT,QTYJUAL_OBAT,JMLHRG_NETTO from TBHISDETILJUAL where BLNTHN_RESEP = \'{1}{2}\''.format(temp_dir + os.path.sep, last_month, last_year))

  lrmf = load_raw_master('{0}TBMASOBAT.TXT'.format(temp_dir + os.path.sep))
  patt_rm = get_pattern(lrmf)
  list_master = create_master_list(lrmf, patt_rm)

  
  list_nama_shaft = {'DISPLAY':[], 'NON-DISPLAY':[]}
  with open('nama_shaft.txt','r') as ns:
    for line in ns:
      if string.split(re.sub('\n','',line),'|')[1] == 'DISPLAY':
        list_nama_shaft['DISPLAY'].append(string.split(re.sub('\n','',line),'|')[0])
      else:
        list_nama_shaft['NON-DISPLAY'].append(string.split(re.sub('\n','',line),'|')[0])

    
  list_master_corrected = {}
  for i in list_master:
    qty_akh = 0
    sign = i[4][0]
    num = float(i[4][1:17])
    exp = int(i[4][-3:])
    
    if sign == '+':
      qty_akh += (num * pow(10,exp))
    else:
      qty_akh -= (num * pow(10,exp))
    
    if (i[9] != "-") and (i[10] != "-"):
      tgl1 = int(i[9][3:5])
      bln1 = int(i[9][:2])
      thn1 = int(i[9][-4:])
      tanggal1 = time.mktime((thn1, bln1, tgl1, 0, 0, 0, 0, 0, 0))
      tgl2 = int(i[10][3:5])
      bln2 = int(i[10][:2])
      thn2 = int(i[10][-4:])
      tanggal2 = time.mktime((thn2, bln2, tgl2, 0, 0, 0, 0, 0, 0))
    elif i[9] == "-" or i[10] == "-":
      tanggal1 = 0
      tanggal2 = 0
      
																			      #loc  sat     qty             hnappn                             faktor           hjah          disc              
    list_master_corrected[i[0]] = [i[1], 'DISPLAY' if i[2] in list_nama_shaft['DISPLAY'] else 'NON-DISPLAY' if i[2] in list_nama_shaft['NON-DISPLAY'] else '',i[2], i[3], qty_akh, Decimal('{0}'.format(i[5])), Decimal('{0}'.format(i[6])), int(i[7][:-3]), Decimal(i[8]), tanggal1, tanggal2]
    ## [KODE_OBAT] = [NAMA_OBAT,JENIS,LOKASI,SATUAN_OBAT,QTYAKHIR_OBAT,HNAPPNREG_OBAT,FAKTOR3,HJAH,DISCH,TANGGAL1,TANGGAL2]
    
    
  data_qty_min_display = {}
  # KODE_OBAT, NAMA_OBAT, LOKASI, SATUAN_OBAT, QTY_MIN_DISPLAY
  with open('qty_display.txt','r') as qd:
    for line in qd:
      if line[0].isdigit():
        data_qty_min_display[string.split(line,'\t')[0]] = int(string.split(line,'\t')[-1])
  
  list_qty_min_display = {}
  #for i in 
  
 
  ltbhdj1 = load_raw_master('{0}TBHISDETILJUAL_1.TXT'.format(temp_dir + os.path.sep))
  patt_tbhdj1 = get_pattern(ltbhdj1)
  list_tbhisdetiljual1 = create_master_list(ltbhdj1, patt_tbhdj1)
  
  ltbhdj2 = load_raw_master('{0}TBHISDETILJUAL_2.TXT'.format(temp_dir + os.path.sep))
  patt_tbhdj2 = get_pattern(ltbhdj2)
  list_tbhisdetiljual2 = create_master_list(ltbhdj2, patt_tbhdj2)
  
  ltbhdj3 = load_raw_master('{0}TBHISDETILJUAL_3.TXT'.format(temp_dir + os.path.sep))
  patt_tbhdj3 = get_pattern(ltbhdj3)
  list_tbhisdetiljual3 = create_master_list(ltbhdj3, patt_tbhdj3)
  
  list_tbhisdetiljual_gabung = [list_tbhisdetiljual1, list_tbhisdetiljual2, list_tbhisdetiljual3]
  
  
  dict_tbhisdetiljual1 = {}
  for i in list_tbhisdetiljual1:
    if i[0] not in dict_tbhisdetiljual1.keys():
      dict_tbhisdetiljual1[i[0]] = {}
      dict_tbhisdetiljual1[i[0]]['NAMA_OBAT'] = i[1]
      dict_tbhisdetiljual1[i[0]]['QTYJUAL_OBAT'] = [int(i[2][:-5])]
      dict_tbhisdetiljual1[i[0]]['JMLHRG_NETTO'] = [int(i[3][:-5])]
    else:
      dict_tbhisdetiljual1[i[0]]['QTYJUAL_OBAT'].append(int(i[2][:-5]))
      dict_tbhisdetiljual1[i[0]]['JMLHRG_NETTO'].append(int(i[3][:-5]))
  
  dict_tbhisdetiljual2 = {}
  for i in list_tbhisdetiljual2:
    if i[0] not in dict_tbhisdetiljual2.keys():
      dict_tbhisdetiljual2[i[0]] = {}
      dict_tbhisdetiljual2[i[0]]['NAMA_OBAT'] = i[1]
      dict_tbhisdetiljual2[i[0]]['QTYJUAL_OBAT'] = [int(i[2][:-5])]
      dict_tbhisdetiljual2[i[0]]['JMLHRG_NETTO'] = [int(i[3][:-5])]
    else:
      dict_tbhisdetiljual2[i[0]]['QTYJUAL_OBAT'].append(int(i[2][:-5]))
      dict_tbhisdetiljual2[i[0]]['JMLHRG_NETTO'].append(int(i[3][:-5]))
  
  dict_tbhisdetiljual3 = {}
  for i in list_tbhisdetiljual3:
    if i[0] not in dict_tbhisdetiljual3.keys():
      dict_tbhisdetiljual3[i[0]] = {}
      dict_tbhisdetiljual3[i[0]]['NAMA_OBAT'] = i[1]
      dict_tbhisdetiljual3[i[0]]['QTYJUAL_OBAT'] = [int(i[2][:-5])]
      dict_tbhisdetiljual3[i[0]]['JMLHRG_NETTO'] = [int(i[3][:-5])]
    else:
      dict_tbhisdetiljual3[i[0]]['QTYJUAL_OBAT'].append(int(i[2][:-5]))
      dict_tbhisdetiljual3[i[0]]['JMLHRG_NETTO'].append(int(i[3][:-5]))
      
  
  dict_tbhisdetiljual_gabung = [dict_tbhisdetiljual1, dict_tbhisdetiljual2, dict_tbhisdetiljual3]


  list_pareto_nilai1 = []
  for i in dict_tbhisdetiljual1.keys():
    list_pareto_nilai1.append([i, dict_tbhisdetiljual1[i]['NAMA_OBAT'], sum(dict_tbhisdetiljual1[i]['JMLHRG_NETTO'])])
  list_pareto_nilai1 = sorted(list_pareto_nilai1, key=itemgetter(2), reverse = True)

  list_pareto_nilai2 = []
  for i in dict_tbhisdetiljual2.keys():
    list_pareto_nilai2.append([i, dict_tbhisdetiljual2[i]['NAMA_OBAT'], sum(dict_tbhisdetiljual2[i]['JMLHRG_NETTO'])])
  list_pareto_nilai2 = sorted(list_pareto_nilai2, key=itemgetter(2), reverse = True)
  
  list_pareto_nilai3 = []
  for i in dict_tbhisdetiljual3.keys():
    list_pareto_nilai3.append([i, dict_tbhisdetiljual3[i]['NAMA_OBAT'], sum(dict_tbhisdetiljual3[i]['JMLHRG_NETTO'])])
  list_pareto_nilai3 = sorted(list_pareto_nilai3, key=itemgetter(2), reverse = True)
  
  list_pareto_kunj1 = []
  for i in dict_tbhisdetiljual1.keys():
    list_pareto_kunj1.append([i, dict_tbhisdetiljual1[i]['NAMA_OBAT'], len(dict_tbhisdetiljual1[i]['QTYJUAL_OBAT'])])
  list_pareto_kunj1 = sorted(list_pareto_kunj1, key=itemgetter(2), reverse = True)
  
  list_pareto_kunj2 = []
  for i in dict_tbhisdetiljual2.keys():
    list_pareto_kunj2.append([i, dict_tbhisdetiljual2[i]['NAMA_OBAT'], len(dict_tbhisdetiljual2[i]['QTYJUAL_OBAT'])])
  list_pareto_kunj2 = sorted(list_pareto_kunj2, key=itemgetter(2), reverse = True)
  
  list_pareto_kunj3 = []
  for i in dict_tbhisdetiljual3.keys():
    list_pareto_kunj3.append([i, dict_tbhisdetiljual3[i]['NAMA_OBAT'], len(dict_tbhisdetiljual3[i]['QTYJUAL_OBAT'])])
  list_pareto_kunj3 = sorted(list_pareto_kunj3, key=itemgetter(2), reverse = True)

  
  list_sorted = [list_pareto_nilai1, list_pareto_nilai2, list_pareto_nilai3, list_pareto_kunj1, list_pareto_kunj2, list_pareto_kunj3]

  
  list_big80 = []
  for i in list_sorted:
    list_big80.append(get_big_80(i))

  
  list_kode_obat_in_pareto = []
  for i in list_sorted[0:3]:
    for j in i:
      if j[0] not in list_kode_obat_in_pareto:
        list_kode_obat_in_pareto.append(j[0])

  
  list_kode_obat_aktif = []
  for i in list_master_corrected.keys():
    if i in list_kode_obat_in_pareto or list_master_corrected[i][2] in list_nama_shaft['DISPLAY'] + list_nama_shaft['NON-DISPLAY'] or list_master_corrected[i][4] != 0:
      list_kode_obat_aktif.append(i)

  
  qty_cure = {}
  for i in list_kode_obat_aktif:
    qty_per_cure = []
    if i in dict_tbhisdetiljual1.keys():
      qty_per_cure += dict_tbhisdetiljual1[i]['QTYJUAL_OBAT']
    if i in dict_tbhisdetiljual2.keys():
      qty_per_cure += dict_tbhisdetiljual2[i]['QTYJUAL_OBAT']
    if i in dict_tbhisdetiljual3.keys():
      qty_per_cure += dict_tbhisdetiljual3[i]['QTYJUAL_OBAT']
    
    if len(qty_per_cure) > 0:
      freq_qty = {}
      for j in qty_per_cure:
        if j not in freq_qty.keys():
          freq_qty[j] = 1
        else:
          freq_qty[j] += 1
    
      find_the_most_frequent = []
      for k in freq_qty.keys():
        if len(find_the_most_frequent) == 0:
          find_the_most_frequent.append(k)
        else:
          if freq_qty[k] > freq_qty[find_the_most_frequent[-1]]:
            find_the_most_frequent.append(k)
      qty_cure[i] = find_the_most_frequent[-1]
    else:
      qty_cure[i] = 0

  
  gabung_80_qty_nilai1 = []
  for i in (list_big80[0], list_big80[3]):
    for j in i:
      if j[0] not in gabung_80_qty_nilai1:
        gabung_80_qty_nilai1.append(j[0])
  
  gabung_80_qty_nilai2 = []
  for i in (list_big80[1], list_big80[2]):
    for j in i:
      if j[0] not in gabung_80_qty_nilai2:
        gabung_80_qty_nilai2.append(j[0])
  
  gabung_80_qty_nilai3 = []
  for i in (list_big80[3], list_big80[3]):
    for j in i:
      if j[0] not in gabung_80_qty_nilai3:
        gabung_80_qty_nilai3.append(j[0])
  
  more_than_once_gabung_80_qty_nilai_123 = []
  for i in list_kode_obat_aktif:
    skor = 0
    for j in (gabung_80_qty_nilai1, gabung_80_qty_nilai2, gabung_80_qty_nilai3):
      if i in j:
        skor += 1
    if skor > 1:
      more_than_once_gabung_80_qty_nilai_123.append(i)


  list_kode_obat_pareto_nilai1 = []
  for i in list_big80[0]:
    list_kode_obat_pareto_nilai1.append(i[0])
    
  list_kode_obat_pareto_kunj1 = []
  for i in list_big80[3]:
    list_kode_obat_pareto_kunj1.append(i[0])

  irisan_80_qty_nilai1 = []
  for i in gabung_80_qty_nilai1:
    if i in list_kode_obat_pareto_nilai1 and i in list_kode_obat_pareto_kunj1:
      irisan_80_qty_nilai1.append(i)                   # ==========> Pareto A di bulan 1


  list_kode_obat_pareto_nilai2 = []
  for i in list_big80[1]:
    list_kode_obat_pareto_nilai2.append(i[0])
    
  list_kode_obat_pareto_kunj2 = []
  for i in list_big80[4]:
    list_kode_obat_pareto_kunj2.append(i[0])
  
  irisan_80_qty_nilai2 = []
  for i in gabung_80_qty_nilai2:
    if i in list_kode_obat_pareto_nilai2 and i in list_kode_obat_pareto_kunj2:
      irisan_80_qty_nilai2.append(i)                   # ==========> Pareto A di bulan 2


  list_kode_obat_pareto_nilai3 = []
  for i in list_big80[2]:
    list_kode_obat_pareto_nilai3.append(i[0])
    
  list_kode_obat_pareto_kunj3 = []
  for i in list_big80[5]:
    list_kode_obat_pareto_kunj3.append(i[0])

  irisan_80_qty_nilai3 = []
  for i in gabung_80_qty_nilai3:
    if i in list_kode_obat_pareto_nilai3 and i in list_kode_obat_pareto_kunj3:
      irisan_80_qty_nilai3.append(i)                    # ==========> Pareto A di bulan 3


  list_kode_obat_dokter_inhouse = []
  with open('obat_dokter_inhouse.txt','r') as di:
    for line in di:
      if line[0].isdigit():
        list_kode_obat_dokter_inhouse.append(re.sub("^ *| *$","",string.split(line,'\t')[0]))


  pareto_a = []
  for i in more_than_once_gabung_80_qty_nilai_123:
    skor = 0
    if i in irisan_80_qty_nilai1:
      skor += 1
    if i in irisan_80_qty_nilai2:
      skor += 1
    if i in irisan_80_qty_nilai3:
      skor += 1
    if skor > 1:
      pareto_a.append(i)                                   # ==========> PARETO A
  for i in list_kode_obat_dokter_inhouse:
    pareto_a.append(i)

  
  pareto_b = []
  for i in more_than_once_gabung_80_qty_nilai_123:
    if i not in pareto_a:
      pareto_b.append(i)                                   # ==========> PARETO B
  
  
  pareto_c = []
  for i in list_kode_obat_aktif:
    if i not in more_than_once_gabung_80_qty_nilai_123:
      pareto_c.append(i)                                   # ==========> PARETO C
  

  ll = [list_master,list_master_corrected,list_qty_min_display,list_tbhisdetiljual1,list_tbhisdetiljual2,list_tbhisdetiljual3,dict_tbhisdetiljual1,dict_tbhisdetiljual2,dict_tbhisdetiljual3,list_pareto_nilai1,list_pareto_nilai2,list_pareto_nilai3,list_pareto_kunj1,list_pareto_kunj2,list_pareto_kunj3,list_big80[0],list_big80[1],list_big80[2],list_big80[3],list_big80[4],list_big80[5],list_kode_obat_in_pareto,list_kode_obat_aktif,qty_cure,gabung_80_qty_nilai1,gabung_80_qty_nilai2,gabung_80_qty_nilai3,more_than_once_gabung_80_qty_nilai_123,list_kode_obat_pareto_nilai1,list_kode_obat_pareto_kunj1,irisan_80_qty_nilai1,list_kode_obat_pareto_nilai2,list_kode_obat_pareto_kunj2,irisan_80_qty_nilai2,list_kode_obat_pareto_nilai3,list_kode_obat_pareto_kunj3,irisan_80_qty_nilai3,pareto_a,pareto_b,pareto_c]
  lt = ['list_master','list_master_corrected','list_qty_min_display','list_tbhisdetiljual1','list_tbhisdetiljual2','list_tbhisdetiljual3','dict_tbhisdetiljual1','dict_tbhisdetiljual2','dict_tbhisdetiljual3','list_pareto_nilai1','list_pareto_nilai2','list_pareto_nilai3','list_pareto_kunj1','list_pareto_kunj2','list_pareto_kunj3','list_big80[0]','list_big80[1]','list_big80[2]','list_big80[3]','list_big80[4]','list_big80[5]','list_kode_obat_in_pareto','list_kode_obat_aktif','qty_cure','gabung_80_qty_nilai1','gabung_80_qty_nilai2','gabung_80_qty_nilai3','more_than_once_gabung_80_qty_nilai_123','list_kode_obat_pareto_nilai1','list_kode_obat_pareto_kunj1','irisan_80_qty_nilai1','list_kode_obat_pareto_nilai2','list_kode_obat_pareto_kunj2','irisan_80_qty_nilai2','list_kode_obat_pareto_nilai3','list_kode_obat_pareto_kunj3','irisan_80_qty_nilai3','pareto_a','pareto_b','pareto_c']
  dl = {}
  for i in range(len(lt)):
    dl[lt[i]] = ll[i]
  for i in lt:
    print("{0}: {1}".format(i, len(dl[i])))
  
  #with open('master.txt','w') as mas:
    #for i in list_kode_obat_aktif:
      #mas.write("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}\t{10}\t{11}\t{12}\n".format(i,list_master_corrected[i][0],list_master_corrected[i][1],list_master_corrected[i][2],list_master_corrected[i][3],list_master_corrected[i][4],list_master_corrected[i][5],list_master_corrected[i][6],list_master_corrected[i][7],list_master_corrected[i][8],list_master_corrected[i][9],list_master_corrected[i][10],int(time.time())))

  monthly_max_usage = {}
  for i in list_kode_obat_aktif:
    first = sum(dict_tbhisdetiljual1[i]['QTYJUAL_OBAT']) if i in dict_tbhisdetiljual1.keys() else 0
    second = sum(dict_tbhisdetiljual2[i]['QTYJUAL_OBAT']) if i in dict_tbhisdetiljual2.keys() else 0
    third = sum(dict_tbhisdetiljual3[i]['QTYJUAL_OBAT']) if i in dict_tbhisdetiljual3.keys() else 0
    monthly_max_usage[i] = max(first, second, third)

  faktor_pengali_pareto = {"A":Decimal('1.0'), "B":Decimal('0.75'), "C":Decimal('0.5')}

  target_stock_based_on_pareto = {}
  for i in list_kode_obat_aktif:
    this_qty_cure = qty_cure[i]
    
    if i in pareto_a:
      qty_temp = int(math.ceil(monthly_max_usage[i] * faktor_pengali_pareto["A"]))
    elif i in pareto_b:
      qty_temp = int(math.ceil(monthly_max_usage[i] * faktor_pengali_pareto["B"]))
    elif i in pareto_c:
      qty_temp = int(math.ceil(monthly_max_usage[i] * faktor_pengali_pareto["C"]))
    
    if this_qty_cure != 0 and qty_temp % this_qty_cure != 0:
      target_stock_based_on_pareto[i] = qty_temp + (this_qty_cure - (qty_temp % this_qty_cure))
    else:
      target_stock_based_on_pareto[i] = qty_temp
  
  #with open('target_stock_based_on_pareto.txt','w') as tsp:
    #for i in target_stock_based_on_pareto.keys():
      #tsp.write('{0}\t{1}\n'.format(list_master_corrected[i][0], target_stock_based_on_pareto[i]))
  
  max_of_them = {}
  for i in list_kode_obat_aktif:
    target_qty_pareto = target_stock_based_on_pareto[i]
    this_qty_cure_2 = qty_cure[i]
    this_qty_min_disp = data_qty_min_display[i] if i in data_qty_min_display.keys() else 0
    max_of_them[i] = max(target_qty_pareto, this_qty_cure_2, this_qty_min_disp)
  
 
  # INPUT : UPDATE_HARGA_ROP-OLD.TXT
  # JENIS_BARANG | LOKASI | KODE_OBAT | NAMA_OBAT | SATUAN | HARGA | DISKON | ROP | KET
  
  old_data = {}
  with open('UPDATE_HARGA_ROP-OLD.TXT','r') as old:
    for line in old:
      if string.split(line, '\t')[2][0].isdigit():
        the_key = string.split(line, '\t')[2]
        old_data[the_key] = [the_key, string.split(line, '\t')[3], string.split(line, '\t')[4], int(string.split(line, '\t')[5]) if string.split(line, '\t')[5] != '' else '', Decimal(string.split(line, '\t')[6]) if string.split(line, '\t')[6] != '' else '', int(string.split(line, '\t')[7])]
  
  
  # OUTPUT TARGET : UPDATE_HARGA_ROP-{DATE-TODAY}.TXT
  # JENIS_BARANG | LOKASI | KODE_OBAT | NAMA_OBAT | SATUAN | HARGA | DISKON | ROP | KET
  # [KODE_OBAT] = [NAMA_OBAT,JENIS,LOKASI,SATUAN_OBAT,QTYAKHIR_OBAT,HNAPPNREG_OBAT,FAKTOR3,HJAH,DISCH,TANGGAL1,TANGGAL2]
#----------- 
  with open('UPDATE_HARGA_ROP-{0}{1}{2}.TXT'.format(time.gmtime(time.time()).tm_year, time.gmtime(time.time()).tm_mon, time.gmtime(time.time()).tm_mday),'w') as out_update:
    temp_dict = {}
    for i in list_kode_obat_aktif:
      if list_master_corrected[i][1] != 'NON-DISPLAY':
        harga_perkalian_faktor = int(list_master_corrected[i][5] * list_master_corrected[i][6])
        hjah = list_master_corrected[i][7]
        diskon_berlaku = time.time() > list_master_corrected[i][9] and time.time() < list_master_corrected[i][10]
        temp_dict[i] = [list_master_corrected[i][1], list_master_corrected[i][2], i, list_master_corrected[i][0], list_master_corrected[i][3], min(harga_perkalian_faktor, hjah) if hjah != 0 else harga_perkalian_faktor, list_master_corrected[i][8] if diskon_berlaku else 0, max_of_them[i]]
      else:
        temp_dict[i] = [list_master_corrected[i][1], list_master_corrected[i][2], i, list_master_corrected[i][0], list_master_corrected[i][3], '', '', max_of_them[i]]
      #                         0jenis                       1lokasi          2kode          3nama                        4satuan                         5harga                                                                                 6diskon                          7rop
    ket_dict = {}
    for i in old_data.keys():
      if i in temp_dict:
        if temp_dict[i][0] == 'NON-DISPLAY':
          ket = ""
          if old_data[i][1] == temp_dict[i][3] and old_data[i][2] == temp_dict[i][4] and  old_data[i][5] == temp_dict[i][7]:
            ket = "Sama"
          else:
            if old_data[i][1] != temp_dict[i][3]:
              ket += "Nama obat: {0} --> {1}  ".format(old_data[i][1], temp_dict[i][3])
            if old_data[i][2] != temp_dict[i][4]:
              ket += "Satuan: {0} --> {1}  ".format(old_data[i][2], temp_dict[i][4])
            if old_data[i][5] != temp_dict[i][7]:
              ket += "ROP: {0} --> {1}  ".format(old_data[i][5], temp_dict[i][7])
          ket_dict[i] = ket
        else:
          ket = ""
          if old_data[i][1] == temp_dict[i][3] and old_data[i][2] == temp_dict[i][4] and old_data[i][3] == temp_dict[i][5] and old_data[i][4] == temp_dict[i][6] and old_data[i][5] == temp_dict[i][7]:
            ket = "Sama"
          else:
            if old_data[i][1] != temp_dict[i][3]:
              ket += "Nama obat: {0} --> {1}  ".format(old_data[i][1], temp_dict[i][3])
            if old_data[i][2] != temp_dict[i][4]:
              ket += "Satuan: {0} --> {1}  ".format(old_data[i][2], temp_dict[i][4])
            if old_data[i][3] != temp_dict[i][5]:
              ket += "Harga: {0} --> {1}  ".format(old_data[i][3], temp_dict[i][5])
            if old_data[i][4] != temp_dict[i][6]:
              ket += "Diskon: {0} --> {1}  ".format(old_data[i][4], temp_dict[i][6])
            if old_data[i][5] != temp_dict[i][7]:
              ket += "ROP: {0} --> {1}  ".format(old_data[i][5], temp_dict[i][7])
          ket_dict[i] = ket
    for i in list_kode_obat_aktif:
      if i not in ket_dict.keys():
        ket_dict[i] = "Baru"
    
    temp = []
    for i in temp_dict.keys():
      #                  0jenis         1lokasi          2kode            3nama             4satuan                                 5harga                                                      6diskon                                 7rop           8ket
      temp.append([temp_dict[i][0], temp_dict[i][1], temp_dict[i][2], temp_dict[i][3], temp_dict[i][4], temp_dict[i][5] if temp_dict[i][0] != 'NON-DISPLAY' else '', temp_dict[i][6] if temp_dict[i][0] != 'NON-DISPLAY' else '', temp_dict[i][7], ket_dict[i]])
      
    temp2 = sorted(temp, key=itemgetter(0,1,3))
    
    out_update.write("JENIS_BARANG\tLOKASI\tKODE_OBAT\tNAMA_OBAT\tSATUAN\tHARGA\tDISKON\tROP\tKET\n")
    for i in temp2:
      out_update.write("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\n".format(i[0],i[1],i[2],i[3],i[4],i[5],i[6],i[7],i[8]))

#------------------    
  # OUTPUT TARGET : DATA_BARANG_BELUM_MASUK_SHAFT.TXT
  # KODE_OBAT | NAMA_OBAT | SHAFT
  with open('DATA_BARANG_BELUM_MASUK_SHAFT-{0}{1}{2}.TXT'.format(time.gmtime(time.time()).tm_year, time.gmtime(time.time()).tm_mon, time.gmtime(time.time()).tm_mday),'w') as nis:
    temp = []
    for i in list_kode_obat_aktif:
      if list_master_corrected[i][1] == "":
        temp.append([i, list_master_corrected[i][0], list_master_corrected[i][2]])
    temp2 = sorted(temp, key=itemgetter(1))
    nis.write("KODE_OBAT\tNAMA_OBAT\tSHAFT\n")
    for i in temp2:
      nis.write("{0}\t{1}\t{2}\n".format(i[0], i[1], i[2]))
  
  # JENIS_BARANG | LOKASI | KODE_OBAT | NAMA_OBAT | SATUAN | HARGA | DISKON | ROP
