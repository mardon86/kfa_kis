import re, string, os, time, math, random, xlrd, xlwt
from operator import itemgetter

#### OPTIONS #####

LINTAS_ODBC = "y" # y/n
OLD_ODBC = "app2013"
CUR_ODBC = "app2015"
CUR_ODBC_MNTYR = "0115" # MMYY
HITUNG_OBAT_DOKTER = "y" # y/n
HITUNG_QTY_MIN_DISPLAY = "n" # y/n
FAKTOR_PENGALI_PARETO_A = 1.0 # 0.0 - 1.0 ( 0.5 berarti pareto A disediakan untuk setengah bulan )
FAKTOR_PENGALI_PARETO_B = 0.5 # 0.0 - 1.0
FAKTOR_PENGALI_PARETO_C = 0.5 # 0.0 - 1.0


##################

def get_mntyr(bulan_ke, mntyr_pattern):
  sixth_mnth = mntyr_pattern[:2]
  sixth_year = mntyr_pattern[2:]
  mnts = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
  res_mnth = mnts[mnts.index(sixth_mnth) - (6 - bulan_ke)]
  if (mnts.index(sixth_mnth) - (6 - bulan_ke)) < 0:
      res_year = str(int(sixth_year) - 1)
  else:
      res_year = sixth_year
  return res_mnth+res_year


def isolder(mntyr, cur_odbc_mntyr):
    mnth_odbc = int(cur_odbc_mntyr[:2])
    year_odbc = int(cur_odbc_mntyr[2:])
    mnth = int(mntyr[:2])
    year = int(mntyr[2:])
    if year < year_odbc:
        return True
    elif mnth < mnth_odbc:
        return True
    else:
        return False


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


# i, j, k ---> KODE_OBAT, NAMA_OBAT, SORT_ITEM (NILAI atau KUNJ)
def get_big_80(the_list):
  tl = the_list
  sum_values = 0
  for i,j,k in tl:
    sum_values += k
  big_80 = []
  eighty_percent_of_sum_values = 0.8 * sum_values
  sum_big80 = 0
  for i in tl:
    if sum_big80 < eighty_percent_of_sum_values:
      big_80.append(i)
      sum_big80 += i[2]
    else:
      break
  return big_80

    
if __name__ == '__main__':
  
  ## Membuat direktori dengan dirname tanggal dieksekusinya script ini
  res_dir = '{0}{1}{2}'.format(time.gmtime(time.time()).tm_year, str(time.gmtime(time.time()).tm_mon) if len(str(time.gmtime(time.time()).tm_mon)) == 2 else "0"+str(time.gmtime(time.time()).tm_mon), time.gmtime(time.time()).tm_mday if len(str(time.gmtime(time.time()).tm_mday)) == 2 else "0"+str(time.gmtime(time.time()).tm_mday))
  os.mkdir(res_dir)

  
  ## Membuat direktori "HISJUALRAW_POOL", direktori ini digunakan untuk menyimpan histori data penjualan,
  ## jadi kalau data histori penjualan pada bulan tertentu sudah pernah disimpan, gak perlu ngambil data
  ## dari database lagi.
  if os.path.exists('HISJUALRAW_POOL'):
    pass
  else:
    os.mkdir('HISJUALRAW_POOL')
  
  
  ## Menyambungkan dengan odbc 
  os.system('db2 connect to {0}'.format(CUR_ODBC))
  os.system('db2 -z{0}TBMASOBAT.TXT select KODE_OBAT,NAMA_OBAT,KODE_GENOBAT,LOKASI,SATUAN_OBAT,QTYAWAL_OBAT,QTYBELI_OBAT,QTYDROPING_OBAT,QTYJUAL_OBAT,HNAPPNREG_OBAT,FAKTOR3,FAKTOR4,HJAH,HJAA,DISCH,DISCA,TANGGAL1,TANGGAL2 from TBMASOBAT'.format(res_dir + os.path.sep))
  #                                        0         1         2            3      4           5            6            7               8            9              10      11      12   13   14    15    16       17


  ## Mengambil data bulan dan tahun terakhir pada histori penjualan (yang telah divalidasi)
  last_rec = os.popen('db2 select BLNTHN_RESEP from TBHISDETILJUAL order by TANGGAL_RESEP desc fetch first 1 row only').read()[30:34]
  os.system('db2 disconnect {0}'.format(CUR_ODBC))
  

  ## Mengambil data penjualan 6 bulan terakhir, jika data penjualan telah ada, PASS!!
  hispenj = [get_mntyr(1, last_rec), get_mntyr(2, last_rec), get_mntyr(3, last_rec), get_mntyr(4, last_rec), get_mntyr(5, last_rec), last_rec]
  for i in hispenj:
    if os.path.exists('HISJUALRAW_POOL' + os.path.sep + 'TBHISDETILJUAL_{0}.TXT'.format(i)):
      pass
    else:
      if LINTAS_ODBC == "n":
        os.system('db2 connect to {0}'.format(CUR_ODBC))
        os.system('db2 -z{0}TBHISDETILJUAL_{1}.TXT select KODE_OBAT,NAMA_OBAT,QTYJUAL_OBAT,JMLHRG_NETTO from TBHISDETILJUAL where BLNTHN_RESEP = \'{1}\''.format('HISJUALRAW_POOL' + os.path.sep, i))
        os.system('db2 disconnect {0}'.format(CUR_ODBC))
      else:
        if isolder(i, CUR_ODBC_MNTYR):
          os.system('db2 connect to {0}'.format(OLD_ODBC))
          os.system('db2 -z{0}TBHISDETILJUAL_{1}.TXT select KODE_OBAT,NAMA_OBAT,QTYJUAL_OBAT,JMLHRG_NETTO from TBHISDETILJUAL where BLNTHN_RESEP = \'{1}\''.format('HISJUALRAW_POOL' + os.path.sep, i))
          os.system('db2 disconnect {0}'.format(OLD_ODBC))
        else:
          os.system('db2 connect to {0}'.format(CUR_ODBC))
          os.system('db2 -z{0}TBHISDETILJUAL_{1}.TXT select KODE_OBAT,NAMA_OBAT,QTYJUAL_OBAT,JMLHRG_NETTO from TBHISDETILJUAL where BLNTHN_RESEP = \'{1}\''.format('HISJUALRAW_POOL' + os.path.sep, i))
          os.system('db2 disconnect {0}'.format(CUR_ODBC))
  

  ## Mengambil data master obat
  lrmf = load_raw_master('{0}TBMASOBAT.TXT'.format(res_dir + os.path.sep))
  patt_rm = get_pattern(lrmf)
  list_master = create_master_list(lrmf, patt_rm)


  ## Mengambil data shaft pada file "NAMA_SHAFT.xls"
  list_nama_shaft = {'DISPLAY':[], 'NON-DISPLAY':[]}
  nsxl = xlrd.open_workbook('NAMA_SHAFT.xls').sheet_by_index(0)
  for i in range(1,nsxl.nrows):
    if nsxl.cell_value(i,1) == 'DISPLAY':
      list_nama_shaft['DISPLAY'].append(nsxl.cell_value(i,0))
    else:
      list_nama_shaft['NON-DISPLAY'].append(nsxl.cell_value(i,0))


  ## Koreksi: 
  ##   - qty akhir = qty awal + beli - droping - jual
  ##   - Format tanggal awal dan akhir berlakunya diskon.
  ##   - Menambah kolom jenis.
  list_master_corrected = {}
  for i in list_master:
    qty_akh = float(i[5]) + float(i[6]) - float(i[7]) - float(i[8])
        
    if (i[16] != "-") and (i[17] != "-"):
      tgl1 = int(i[16][3:5])
      bln1 = int(i[16][:2])
      thn1 = int(i[16][-4:])
      tanggal1 = time.mktime((thn1, bln1, tgl1, 0, 0, 0, 0, 0, 0))
      tgl2 = int(i[17][3:5])
      bln2 = int(i[17][:2])
      thn2 = int(i[17][-4:])
      tanggal2 = time.mktime((thn2, bln2, tgl2, 0, 0, 0, 0, 0, 0))
    elif i[16] == "-" or i[17] == "-":
      tanggal1 = 0
      tanggal2 = 0

    # KODE_OBAT,NAMA_OBAT,KODE_GENOBAT,LOKASI,SATUAN_OBAT,QTYAWAL_OBAT,QTYBELI_OBAT,QTYDROPING_OBAT,QTYJUAL_OBAT,HNAPPNREG_OBAT,FAKTOR3,FAKTOR4,HJAH,HJAA,DISCH,DISCA,TANGGAL1,TANGGAL2
    # 0         1         2            3      4           5            6            7               8            9              10      11      12   13   14    15    16       17

    list_master_corrected[i[0]] = [i[1], i[2], 'DISPLAY' if i[3] in list_nama_shaft['DISPLAY'] else 'NON-DISPLAY' if i[3] in list_nama_shaft['NON-DISPLAY'] else '', i[3], i[4], qty_akh, float('{0}'.format(i[9])), float('{0}'.format(i[10])), float('{0}'.format(i[11])), int(i[12][:-3]), int(i[13][:-3]), float(i[14]), float(i[15]), tanggal1, tanggal2]
    ##              [KODE_OBAT] = [NAMA_OBAT,KODE_GENOBAT,JENIS,                                                                                                     LOKASI,SATUAN_OBAT,QTYAKHIR_OBAT,HNAPPNREG_OBAT,FAKTOR3,                    FAKTOR4,                    HJAH,            HJAA,            DISCH,        DISCA,        TANGGAL1, TANGGAL2]
    ##               0             1         2            C                                                                                                          3      4           C             9              10                          11                          12               13               14            15            C         C
    
  
  ## Mengambil data quantity minimal display  
  data_qty_min_display = {}
  # KODE_OBAT, NAMA_OBAT, LOKASI, SATUAN_OBAT, QTY_MIN_DISPLAY
  if HITUNG_QTY_MIN_DISPLAY == "y":
    qdxl = xlrd.open_workbook('QTY_DISPLAY.xls').sheet_by_index(0)
    for i in range(1,qdxl.nrows):
      data_qty_min_display[qdxl.cell_value(i,0)] = qdxl.cell_value(i,4)
  

  ## Membuat dictionary gabungan tbhisdetiljual selama 6 bulan
  ## Dictionary of Lists
  list_tbhisdetiljual_gabung = {}
  for i,j in enumerate(hispenj):
    ltbhdj = load_raw_master('{0}TBHISDETILJUAL_{1}.TXT'.format('HISJUALRAW_POOL' + os.path.sep, j))
    patt_tbhdj = get_pattern(ltbhdj)
    list_tbhisdetiljual_gabung[i+1] = create_master_list(ltbhdj, patt_tbhdj)

  
  ## Membuat dictionary dari tbhisdetiljual selama 6 bulan yang di buat dictionary dengan kode obat sebagai key nya.
  ## Dictionary of dictionaries
  ## QTYJUAL_OBAT = [..,...,...,..]
  ## JMLHRG_NETTO = [..,...,...,..]
  dict_tbhisdetiljual_gabung = {1:{}, 2:{}, 3:{}, 4:{}, 5:{}, 6:{}}
  for month in range(1,7):
    for row in list_tbhisdetiljual_gabung[month]:
      if row[0] not in dict_tbhisdetiljual_gabung[month].keys():
        kode_obat = row[0]
        dict_tbhisdetiljual_gabung[month][kode_obat] = {}
        dict_tbhisdetiljual_gabung[month][kode_obat]['NAMA_OBAT'] = row[1]
        dict_tbhisdetiljual_gabung[month][kode_obat]['QTYJUAL_OBAT'] = [int(row[2][:-5])]
        dict_tbhisdetiljual_gabung[month][kode_obat]['JMLHRG_NETTO'] = [int(row[3][:-5])]
      else:
        dict_tbhisdetiljual_gabung[month][kode_obat]['QTYJUAL_OBAT'].append(int(row[2][:-5]))
        dict_tbhisdetiljual_gabung[month][kode_obat]['JMLHRG_NETTO'].append(int(row[3][:-5]))

  
  ## Membuat List gabungan penjualan selama 3 bulan terakhir, diurutkan berdasarkan nilai penjualan tertinggi
  list_pareto_nilai = {}
  for month in range(4,7):
    list_pareto_nilai[month] = []
    for kode_obat in dict_tbhisdetiljual_gabung[month].keys():
      list_pareto_nilai[month].append([kode_obat, dict_tbhisdetiljual_gabung[month][kode_obat]['NAMA_OBAT'], sum(dict_tbhisdetiljual_gabung[month][kode_obat]['JMLHRG_NETTO'])])
    list_pareto_nilai[month] = sorted(list_pareto_nilai[month], key=itemgetter(2), reverse = True)


  ## Membuat List gabungan penjualan selama 3 bulan terakhir, diurutkan berdasarkan jumlah kunjungan tertinggi
  list_pareto_kunj = {}
  for month in range(4,7):
    list_pareto_kunj[month] = []
    for kode_obat in dict_tbhisdetiljual_gabung[month].keys():
      list_pareto_kunj[month].append([kode_obat, dict_tbhisdetiljual_gabung[month][kode_obat]['NAMA_OBAT'], len(dict_tbhisdetiljual_gabung[month][kode_obat]['QTYJUAL_OBAT'])])
    list_pareto_kunj[month] = sorted(list_pareto_kunj[month], key=itemgetter(2), reverse = True)
  

  ## Mengambil 80% nilai tertinggi dari list gabungan penjualan selama 3 bulan terakhir  
  list_big80_pareto_nilai = []
  for month in list_pareto_nilai.keys():
    list_big80_pareto_nilai.append(get_big_80(list_pareto_nilai[month]))


  ## Mengambil 80% kunjungan tertinggi dari list gabungan penjualan selama 3 bulan terakhir
  list_big80_pareto_kunj = []
  for month in list_pareto_kunj.keys():
    list_big80_pareto_kunj.append(get_big_80(list_pareto_kunj[month]))
  

  ## Membuat list kode obat yang ada dalam pareto (ada penjualan dalam 3 bulan terakhir)
  list_kode_obat_in_pareto = []
  for month in list_pareto_nilai.keys():
    for j in list_pareto_nilai[month]:
      if j[0] not in list_kode_obat_in_pareto:
        list_kode_obat_in_pareto.append(j[0])

  ################################# BOOKMARK #################################################
  ## Membuat list kode obat aktif
  list_kode_obat_aktif = []
  for i in list_master_corrected.keys():
    if i in list_kode_obat_in_pareto or list_master_corrected[i][2] in list_nama_shaft['DISPLAY'] + list_nama_shaft['NON-DISPLAY'] or list_master_corrected[i][5] != 0:
      list_kode_obat_aktif.append(i)


  qty_cure = {}
  for kode_obat in list_kode_obat_aktif:
    qty_per_cure = []
    for month in range(4,7):
      if kode_obat in dict_tbhisdetiljual_gabung[month].keys():
        qty_per_cure += dict_tbhisdetiljual_gabung[month]['QTYJUAL_OBAT']
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
      qty_cure[kode_obat] = find_the_most_frequent[-1]
    else:
      qty_cure[kode_obat] = 0

  
  gabung_80_nilai_kunj = {}
  for month in range(4,7):
    gabung_80_nilai_kunj[month] = []
    for i in (list_big80_pareto_nilai[0], list_big80_pareto_kunj[0]):
      for j in i:
        if j[0] not in gabung_80_nilai_kunj[month]:
          gabung_80_nilai_kunj[month].append(j[0])
  
    
  more_than_once_gabung_80_nilai_kunj_123 = []
  for i in list_kode_obat_aktif:
    skor = 0
    for j in (gabung_80_nilai_kunj[4], gabung_80_nilai_kunj[5], gabung_80_nilai_kunj[6]):
      if i in j:
        skor += 1
    if skor > 1:
      more_than_once_gabung_80_nilai_kunj_123.append(i)


  irisan_80_nilai_kunj = {}  # ==========> Pareto A di bulan 1, 2, dan 3

  for month in range(4,7):
    list_kode_obat_pareto_nilai = []
    for i in list_big80_pareto_nilai[month]:
      list_kode_obat_pareto_nilai.append(i[0])
    
    list_kode_obat_pareto_kunj = []
    for i in list_big80_pareto_kunj[month]:
      list_kode_obat_pareto_kunj.append(i[0])

    irisan_80_nilai_kunj[month] = []
    for i in gabung_80_nilai_kunj:
      if i in list_kode_obat_pareto_nilai and i in list_kode_obat_pareto_kunj:
        irisan_80_nilai_kunj[month].append(i)                   


  list_kode_obat_dokter_inhouse = []
  if HITUNG_OBAT_DOKTER == "y":
    dixl = xlrd.open_workbook('obat_dokter_inhouse.xls').sheet_by_index(0)
    for i in range(1, dixl.nrows):
      list_kode_obat_dokter_inhouse.append(dixl.cell_value(i,0))


  pareto_a = []
  for i in more_than_once_gabung_80_nilai_kunj_123:
    skor = 0
    for j in range(4,7):
      if i in irisan_80_nilai_kunj[j]:
        skor += 1
    if skor > 1:
      pareto_a.append(i)                                  # ==========> PARETO A
  
  if HITUNG_OBAT_DOKTER == "y":
    for i in list_kode_obat_dokter_inhouse:
      if i not in pareto_a:
        pareto_a.append(i)

  
  pareto_b = []
  for i in more_than_once_gabung_80_nilai_kunj_123:
    if i not in pareto_a:
      pareto_b.append(i)                                   # ==========> PARETO B
  
  
  pareto_c = []
  for i in list_kode_obat_aktif:
    if i not in more_than_once_gabung_80_nilai_kunj_123:
      pareto_c.append(i)                                   # ==========> PARETO C
  

  ll = [list_master,list_master_corrected,
    list_tbhisdetiljual_gabung[4],list_tbhisdetiljual_gabung[5],list_tbhisdetiljual_gabung[6],
    dict_tbhisdetiljual_gabung[4],dict_tbhisdetiljual_gabung[5],dict_tbhisdetiljual_gabung[6],
    list_pareto_nilai[4],list_pareto_nilai[5],list_pareto_nilai[6],
    list_pareto_kunj[4],list_pareto_kunj[5],list_pareto_kunj[6],
    list_big80_pareto_nilai[0],list_big80_pareto_nilai[1],list_big80_pareto_nilai[2],
    list_big80_pareto_kunj[0],list_big80_pareto_kunj[1],list_big80_pareto_kunj[2],
    list_kode_obat_in_pareto,list_kode_obat_aktif,qty_cure,
    gabung_80_qty_nilai[4],gabung_80_qty_nilai[5],gabung_80_qty_nilai[6],
    more_than_once_gabung_80_qty_nilai_123,
    list_kode_obat_pareto_nilai[4],list_kode_obat_pareto_kunj[4],irisan_80_qty_nilai[4],
    list_kode_obat_pareto_nilai[5],list_kode_obat_pareto_kunj[5],irisan_80_qty_nilai[5],
    list_kode_obat_pareto_nilai[6],list_kode_obat_pareto_kunj[6],irisan_80_qty_nilai[6],
    pareto_a,pareto_b,pareto_c]
  
  lt = ['list_master','list_master_corrected',
    'list_tbhisdetiljual1','list_tbhisdetiljual2','list_tbhisdetiljual3',
    'dict_tbhisdetiljual1','dict_tbhisdetiljual2','dict_tbhisdetiljual3',
    'list_pareto_nilai1','list_pareto_nilai2','list_pareto_nilai3',
    'list_pareto_kunj1','list_pareto_kunj2','list_pareto_kunj3',
    'list_big80[0]','list_big80[1]','list_big80[2]',
    'list_big80[3]','list_big80[4]','list_big80[5]',
    'list_kode_obat_in_pareto','list_kode_obat_aktif','qty_cure',
    'gabung_80_qty_nilai1','gabung_80_qty_nilai2','gabung_80_qty_nilai3',
    'more_than_once_gabung_80_qty_nilai_123',
    'list_kode_obat_pareto_nilai1','list_kode_obat_pareto_kunj1','irisan_80_qty_nilai1',
    'list_kode_obat_pareto_nilai2','list_kode_obat_pareto_kunj2','irisan_80_qty_nilai2',
    'list_kode_obat_pareto_nilai3','list_kode_obat_pareto_kunj3','irisan_80_qty_nilai3',
    'pareto_a','pareto_b','pareto_c']

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

  faktor_pengali_pareto = {"A":FAKTOR_PENGALI_PARETO_A, "B":FAKTOR_PENGALI_PARETO_B, "C":FAKTOR_PENGALI_PARETO_C}

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
  oldxl = xlrd.open_workbook('UPDATE_HARGA_ROP-OLD.xls').sheet_by_index(0)
  for i in range(1, oldxl.nrows):
    the_key = oldxl.cell_value(i,2)
    old_data[the_key] = [the_key, oldxl.cell_value(i,3), oldxl.cell_value(i,4), int(oldxl.cell_value(i,5)) if oldxl.cell_value(i,5) != '' else '', oldxl.cell_value(i,6) if oldxl.cell_value(i,6) != '' else '', int(oldxl.cell_value(i,7))]
  
  
  # OUTPUT TARGET : UPDATE_HARGA_ROP-{DATE-TODAY}.TXT
  # JENIS_BARANG | LOKASI | KODE_OBAT | NAMA_OBAT | SATUAN | HARGA | DISKON | ROP | KET
  # [KODE_OBAT] = [NAMA_OBAT,JENIS,LOKASI,SATUAN_OBAT,QTYAKHIR_OBAT,HNAPPNREG_OBAT,FAKTOR3,HJAH,DISCH,TANGGAL1,TANGGAL2]
#----------- 
  outputxl = xlwt.Workbook()
  outputsh1 = outputxl.add_sheet('Sheet1')
  
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
  
  temp_output = []
  for i in temp_dict.keys():
    #                  0jenis         1lokasi          2kode            3nama             4satuan                                 5harga                                                      6diskon                                 7rop           8ket
    temp_output.append([temp_dict[i][0], temp_dict[i][1], temp_dict[i][2], temp_dict[i][3], temp_dict[i][4], temp_dict[i][5] if temp_dict[i][0] != 'NON-DISPLAY' else '', temp_dict[i][6] if temp_dict[i][0] != 'NON-DISPLAY' else '', temp_dict[i][7], ket_dict[i]])
    
  temp_output2 = sorted(temp_output, key=itemgetter(0,1,3))
  
  clnm = ["JENIS_BARANG", "LOKASI", "KODE_OBAT", "NAMA_OBAT", "SATUAN", "HARGA", "DISKON", "ROP", "KET"]
  
  for i in range(9):
    outputsh1.write(0, i, clnm[i])
  
  for i in range(len(temp_output2)):
    for j in range(len(temp_output2[i])):
      outputsh1.write(i + 1, j, temp_output2[i][j])
  
  outputxl.save('{0}UPDATE_HARGA_ROP-{1}{2}{3}.xls'.format(res_dir + os.path.sep, time.gmtime(time.time()).tm_year, str(time.gmtime(time.time()).tm_mon) if len(str(time.gmtime(time.time()).tm_mon)) == 2 else "0"+str(time.gmtime(time.time()).tm_mon), time.gmtime(time.time()).tm_mday if len(str(time.gmtime(time.time()).tm_mday)) == 2 else "0"+str(time.gmtime(time.time()).tm_mday)))
#------------------    


  # OUTPUT TARGET : DATA_BARANG_BELUM_MASUK_SHAFT.TXT
  # KODE_OBAT | NAMA_OBAT | SHAFT
  nisxl = xlwt.Workbook()
  nissh1 = nisxl.add_sheet('Sheet1')
  temp_nis = []
  for i in list_kode_obat_aktif:
    if list_master_corrected[i][1] == "":
      temp_nis.append([i, list_master_corrected[i][0], list_master_corrected[i][2]])
  temp_nis2 = sorted(temp_nis, key=itemgetter(1))
  nissh1.write(0, 0, "KODE_OBAT")
  nissh1.write(0, 1, "NAMA_OBAT")
  nissh1.write(0, 2, "SHAFT")
  for i in range(len(temp_nis2)):
    for j in range(3):
      nissh1.write(i + 1, j, temp_nis2[i][j])
  nisxl.save('{0}DATA_BARANG_BELUM_MASUK_SHAFT-{1}{2}{3}.xls'.format(res_dir + os.path.sep, time.gmtime(time.time()).tm_year, str(time.gmtime(time.time()).tm_mon) if len(str(time.gmtime(time.time()).tm_mon)) == 2 else "0"+str(time.gmtime(time.time()).tm_mon), time.gmtime(time.time()).tm_mday if len(str(time.gmtime(time.time()).tm_mday)) == 2 else "0"+str(time.gmtime(time.time()).tm_mday)))

  # JENIS_BARANG | LOKASI | KODE_OBAT | NAMA_OBAT | SATUAN | HARGA | DISKON | ROP
  
  
  # data_qty_min_display = {}
  # list_kode_obat_aktif
  # list_master_corrected[i[0]] ## [KODE_OBAT] = [NAMA_OBAT,JENIS,LOKASI,SATUAN_OBAT,QTYAKHIR_OBAT,HNAPPNREG_OBAT,FAKTOR3,HJAH,DISCH,TANGGAL1,TANGGAL2]
  
  list_kode_obat_non_nondisplay = []
  for i in list_kode_obat_aktif:
    if list_master_corrected[i][1] != 'NON-DISPLAY':
      list_kode_obat_non_nondisplay.append(i)
  
  list_kode_obat_belum_didata_qty_min_display = []
  for i in list_kode_obat_non_nondisplay:
    if i not in data_qty_min_display.keys():
      list_kode_obat_belum_didata_qty_min_display.append(i)
  
  lbdqmdxl = xlwt.Workbook()
  lbdqmdsh1 = lbdqmdxl.add_sheet('Sheet1')
  #clnm_lbdqmd = ["LOKASI", "KODE_OBAT", "NAMA_OBAT", "SATUAN", "QTY_MIN_DISPLAY"]
  clnm_lbdqmd = ["KODE_OBAT", "NAMA_OBAT", "LOKASI", "SATUAN_OBAT", "QTY_MIN_DISPLAY"]
  for i in range(len(clnm_lbdqmd)):
    lbdqmdsh1.write(0, i, clnm_lbdqmd[i])
  list_lbdqmd = []
  for i in range(len(list_kode_obat_belum_didata_qty_min_display)):
    list_lbdqmd.append([list_kode_obat_belum_didata_qty_min_display[i], list_master_corrected[list_kode_obat_belum_didata_qty_min_display[i]][0], list_master_corrected[list_kode_obat_belum_didata_qty_min_display[i]][2], list_master_corrected[list_kode_obat_belum_didata_qty_min_display[i]][3]])
  list_sorted_lbdqmd = sorted(list_lbdqmd, key=itemgetter(2,1))
  for i in range(len(list_sorted_lbdqmd)):
    lbdqmdsh1.write(i + 1, 0, list_sorted_lbdqmd[i][0])
    lbdqmdsh1.write(i + 1, 1, list_sorted_lbdqmd[i][1])
    lbdqmdsh1.write(i + 1, 2, list_sorted_lbdqmd[i][2])
    lbdqmdsh1.write(i + 1, 3, list_sorted_lbdqmd[i][3])
    lbdqmdsh1.write(i + 1, 4, "")
  lbdqmdxl.save('{0}QTY_MIN_DISPLAY_YANG_BELUM_DIDATA-{1}{2}{3}.xls'.format(res_dir + os.path.sep, time.gmtime(time.time()).tm_year, str(time.gmtime(time.time()).tm_mon) if len(str(time.gmtime(time.time()).tm_mon)) == 2 else "0"+str(time.gmtime(time.time()).tm_mon), time.gmtime(time.time()).tm_mday if len(str(time.gmtime(time.time()).tm_mday)) == 2 else "0"+str(time.gmtime(time.time()).tm_mday)))


  pareto_not_in_master = []

 
  dict_pareto_nilai1 = {}
  for i,j,k in list_pareto_nilai1:
    dict_pareto_nilai1[i] = [j, k]
  
  dict_pareto_nilai2 = {}
  for i,j,k in list_pareto_nilai2:
    dict_pareto_nilai2[i] = [j, k]
  
  dict_pareto_nilai3 = {}
  for i,j,k in list_pareto_nilai3:
    dict_pareto_nilai3[i] = [j, k]
  
  dict_pareto_kunj1 = {}
  for i,j,k in list_pareto_kunj1:
    dict_pareto_kunj1[i] = [j, k]
  
  dict_pareto_kunj2 = {}
  for i,j,k in list_pareto_kunj2:
    dict_pareto_kunj2[i] = [j, k]
  
  dict_pareto_kunj3 = {}
  for i,j,k in list_pareto_kunj3:
    dict_pareto_kunj3[i] = [j, k]

  clnm_pareto = ["KODE_OBAT", "NAMA_OBAT", "JENIS", "LOKASI", "SATUAN_OBAT", "STOK", "PENJ_1", "KUNJ_1", "PENJ_2", "KUNJ_2", "PENJ_3", "KUNJ_3"]
  
  par_a_xl = xlwt.Workbook()
  par_a_xlsh1 = par_a_xl.add_sheet('Sheet1')
  for i in range(len(clnm_pareto)):
    par_a_xlsh1.write(0, i, clnm_pareto[i])
  lpar_a = []
  for i in pareto_a:
    if i in list_master_corrected.keys():
      lpar_a.append([i, list_master_corrected[i][0], list_master_corrected[i][1], list_master_corrected[i][2], list_master_corrected[i][3], list_master_corrected[i][4], dict_pareto_nilai1[i][1] if i in dict_pareto_nilai1.keys() else 0, dict_pareto_kunj1[i][1]  if i in dict_pareto_kunj1.keys() else 0, dict_pareto_nilai2[i][1] if i in dict_pareto_nilai2.keys() else 0, dict_pareto_kunj2[i][1] if i in dict_pareto_kunj2.keys() else 0, dict_pareto_nilai3[i][1] if i in dict_pareto_nilai3.keys() else 0, dict_pareto_kunj3[i][1] if i in dict_pareto_kunj3.keys() else 0])
    else:
      pareto_not_in_master.append([i,'A'])
  lpar_a_sorted = sorted(lpar_a, key=itemgetter(2,3,1))
  for i in range(len(lpar_a_sorted)):
    for j in range(len(clnm_pareto)):
      par_a_xlsh1.write(i + 1, j, lpar_a_sorted[i][j])
  par_a_xl.save('{0}DAFTAR_PARETO_A-{1}{2}{3}.xls'.format(res_dir + os.path.sep, time.gmtime(time.time()).tm_year, str(time.gmtime(time.time()).tm_mon) if len(str(time.gmtime(time.time()).tm_mon)) == 2 else "0"+str(time.gmtime(time.time()).tm_mon), time.gmtime(time.time()).tm_mday if len(str(time.gmtime(time.time()).tm_mday)) == 2 else "0"+str(time.gmtime(time.time()).tm_mday)))
  
  par_b_xl = xlwt.Workbook()
  par_b_xlsh1 = par_b_xl.add_sheet('Sheet1')
  for i in range(len(clnm_pareto)):
    par_b_xlsh1.write(0, i, clnm_pareto[i])
  lpar_b = []
  for i in pareto_b:
    if i in list_master_corrected.keys():
      lpar_b.append([i, list_master_corrected[i][0], list_master_corrected[i][1], list_master_corrected[i][2], list_master_corrected[i][3], list_master_corrected[i][4], dict_pareto_nilai1[i][1] if i in dict_pareto_nilai1.keys() else 0, dict_pareto_kunj1[i][1]  if i in dict_pareto_kunj1.keys() else 0, dict_pareto_nilai2[i][1] if i in dict_pareto_nilai2.keys() else 0, dict_pareto_kunj2[i][1] if i in dict_pareto_kunj2.keys() else 0, dict_pareto_nilai3[i][1] if i in dict_pareto_nilai3.keys() else 0, dict_pareto_kunj3[i][1] if i in dict_pareto_kunj3.keys() else 0])
    else:
      pareto_not_in_master.append([i,'B'])
  lpar_b_sorted = sorted(lpar_b, key=itemgetter(2,3,1))
  for i in range(len(lpar_b_sorted)):
    for j in range(len(clnm_pareto)):
      par_b_xlsh1.write(i + 1, j, lpar_b_sorted[i][j])
  par_b_xl.save('{0}DAFTAR_PARETO_B-{1}{2}{3}.xls'.format(res_dir + os.path.sep, time.gmtime(time.time()).tm_year, str(time.gmtime(time.time()).tm_mon) if len(str(time.gmtime(time.time()).tm_mon)) == 2 else "0"+str(time.gmtime(time.time()).tm_mon), time.gmtime(time.time()).tm_mday if len(str(time.gmtime(time.time()).tm_mday)) == 2 else "0"+str(time.gmtime(time.time()).tm_mday)))
  
  par_c_xl = xlwt.Workbook()
  par_c_xlsh1 = par_c_xl.add_sheet('Sheet1')
  for i in range(len(clnm_pareto)):
    par_c_xlsh1.write(0, i, clnm_pareto[i])
  lpar_c = []
  for i in pareto_c:
    if i in list_master_corrected.keys():
      lpar_c.append([i, list_master_corrected[i][0], list_master_corrected[i][1], list_master_corrected[i][2], list_master_corrected[i][3], list_master_corrected[i][4], dict_pareto_nilai1[i][1] if i in dict_pareto_nilai1.keys() else 0, dict_pareto_kunj1[i][1]  if i in dict_pareto_kunj1.keys() else 0, dict_pareto_nilai2[i][1] if i in dict_pareto_nilai2.keys() else 0, dict_pareto_kunj2[i][1] if i in dict_pareto_kunj2.keys() else 0, dict_pareto_nilai3[i][1] if i in dict_pareto_nilai3.keys() else 0, dict_pareto_kunj3[i][1] if i in dict_pareto_kunj3.keys() else 0])
    else:
      pareto_not_in_master.append([i,'C'])
  lpar_c_sorted = sorted(lpar_c, key=itemgetter(2,3,1))
  for i in range(len(lpar_c_sorted)):
    for j in range(len(clnm_pareto)):
      par_c_xlsh1.write(i + 1, j, lpar_c_sorted[i][j])
  par_c_xl.save('{0}DAFTAR_PARETO_C-{1}{2}{3}.xls'.format(res_dir + os.path.sep, time.gmtime(time.time()).tm_year, str(time.gmtime(time.time()).tm_mon) if len(str(time.gmtime(time.time()).tm_mon)) == 2 else "0"+str(time.gmtime(time.time()).tm_mon), time.gmtime(time.time()).tm_mday if len(str(time.gmtime(time.time()).tm_mday)) == 2 else "0"+str(time.gmtime(time.time()).tm_mday)))
  
  no_master_xl = xlwt.Workbook()
  no_mas_xlsh1 = no_master_xl.add_sheet('Sheet1')
  no_mas_xlsh1.write(0, 0, "KODE_OBAT")
  no_mas_xlsh1.write(0, 1, "PARETO")
  for i in range(len(pareto_not_in_master)):
    no_mas_xlsh1.write(i + 1, 0, pareto_not_in_master[i][0])
    no_mas_xlsh1.write(i + 1, 1, pareto_not_in_master[i][1])
  no_master_xl.save('{0}DAFTAR_BARANG_PARETO_TIDAK_ADA_DIMASTER-{1}{2}{3}.xls'.format(res_dir + os.path.sep, time.gmtime(time.time()).tm_year, str(time.gmtime(time.time()).tm_mon) if len(str(time.gmtime(time.time()).tm_mon)) == 2 else "0"+str(time.gmtime(time.time()).tm_mon), time.gmtime(time.time()).tm_mday if len(str(time.gmtime(time.time()).tm_mday)) == 2 else "0"+str(time.gmtime(time.time()).tm_mday)))
