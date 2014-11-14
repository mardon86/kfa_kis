import re,string,os

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
  amount_of_dashes[-1] =- 1
  return amount_of_dashes


def create_master_list(loaded_raw_master_file, pattern):
  master_list = []
  for line in loaded_raw_master_file:
    if len(line) > 0:
      if line[0] == ' ' or len(line) == 0 or '---' in line:
        continue
      line_list = []
      curs = 0
      for col in pattern:
        line_list.append(re.sub('^ *| *$','',line[curs:curs+col]))
        curs = curs + col + 1
      master_list.append(line_list), calendar
  return master_list

    
if __name__ == '__main__':
# TBMASOBAT  
# KODE_OBAT NAMA_OBAT	LOKASI	SATUAN	STOK	HNAPPNREG_OBAT	FAKTOR3	HJAH	DISCH	TANGGAL1	TANGGAL2	
#     7         8         70       10    40           14           58    67       61       64              65            

# TBHISDETILJUAL
# KODE_OBAT	NAMA_OBAT	QTYJUAL_OBAT	JMLHRG_NETTO
#     8             9                14              20

  temp_dir = '.temp_'+''.join(random.choice('ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789') for _ in range(8))
  os.mkdir(temp_dir)
  
  os.system('db2 connect to app2013')
  os.system('db2 -z{0}TBMASOBAT.TXT select KODE_OBAT,NAMA_OBAT,LOKASI,SATUAN,STOK,HNAPPNREG_OBAT,FAKTOR3,HJAH,DISCH,TANGGAL1,TANGGAL2 from TBMASOBAT'.format(temp_dir + os.path.sep))
  
  last_rec = os.popen('db2 select BLNTHN_RESEP from TBHISDETILJUAL order by TANGGAL_RESEP desc fetch first 1 row only').read()[30:34]
  last_month = last_rec[0:2]
  last_year = last_rec[2:]

  if last_month == "02":
    first_month = "12"
    first_year = "{0}".format(int(last_year) - 1)
    sec_month = "01"
    sec_year = last_yearlist_tbhisdetiljual1
  elif last_month == "01":
    first_month = "11"
    first_year = "{0}".format(int(last_year) - 1)
    sec_month = "12"
    sec_year = "{0}".format(int(last_year) - 1)
  elif last_month == "12":
    first_month = "10"
    first_year = last_year
    sec_month = "11"
    sec year = last_year
  elif last_month == "11":
    first_month = "09"
    first_year = last_year
    sec_month = "10"
    sec_year = last_year
  else:
    first_month = "0{0}".format(int(last_month) - 2)
    first_year = last_year
    sec_month = "0{0}".format(int(last_month) - 1)
    sec_year = last_year

  os.system('db2 -z{0}TBHISDETILJUAL_1.TXT select KODE_OBAT,NAMA_OBAT,QTYJUAL_OBAT,JMLHRG_NETTO from TBHISDETILJUAL where BLNTHN_RESEP = \'{1}{2}\''.format(temp_dir + os.path.sep, first_month, first_year))
  os.system('db2 -z{0}TBHISDETILJUAL_2.TXT select KODE_OBAT,NAMA_OBAT,QTYJUAL_OBAT,JMLHRG_NETTO from TBHISDETILJUAL where BLNTHN_RESEP = \'{1}{2}\''.format(temp_dir + os.path.sep, sec_month, sec_year))
  os.system('db2 -z{0}TBHISDETILJUAL_3.TXT select KODE_OBAT,NAMA_OBAT,QTYJUAL_OBAT,JMLHRG_NETTO from TBHISDETILJUAL where BLNTHN_RESEP = \'{1}{2}\''.format(temp_dir + os.path.sep, last_month, last_year))

  lrmf = load_raw_master('{0}TBMASOBAT.TXT'.format(temp_dir + os.path.sep))
  patt_rm = get_pattern(lrmf)
  list_master = create_master_list(lrmf, patt_rm)
  
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
      dict_tbhisdetiljual1[i[0]]['QTYJUAL_OBAT'] = [int(i[2]))]
      dict_tbhisdetiljual1[i[0]]['JMLHRG_NETTO'] = [int(i[3]]
  
  dict_tbhisdetiljual2 = {}
  for i in list_tbhisdetiljual2:
    if i[0] not in dict_tbhisdetiljual2.keys():
      dict_tbhisdetiljual2[i[0]] = {}
      dict_tbhisdetiljual2[i[0]]['NAMA_OBAT'] = i[1]
      dict_tbhisdetiljual2[i[0]]['QTYJUAL_OBAT'] = [int(i[2])]
      dict_tbhisdetiljual2[i[0]]['JMLHRG_NETTO'] = [int(i[3])]
  
  dict_tbhisdetiljual3 = {}
  for i in list_tbhisdetiljual3:
    if i[0] not in dict_tbhisdetiljual3.keys():
      dict_tbhisdetiljual3[i[0]] = {}
      dict_tbhisdetiljual3[i[0]]['NAMA_OBAT'] = i[1]
      dict_tbhisdetiljual3[i[0]]['QTYJUAL_OBAT'] = [int(i[2])]
      dict_tbhisdetiljual3[i[0]]['JMLHRG_NETTO'] = [int(i[3])]
      
  dict_tbhisdetiljual_gabung = [dict_tbhisdetiljual1, dict_tbhisdetiljual2, dict_tbhisdetiljual3]
  
  class Items:
    def __init__(self,kode_obat):
      
    
  
  list_pareto_nilai1 = []
  for i in dict_tbhisdetiljual1.keys():
    list_pareto_nilai1.append([i, dict_tbhisdetiljual1[i]['NAMA_OBAT'], sum(dict_tbhisdetiljual1[i]['JMLHRG_NETTO'])])
  
  list_pareto_nilai2 = []
  for i in dict_tbhisdetiljual2.keys():
    list_pareto_nilai2.append([i, dict_tbhisdetiljual2[i]['NAMA_OBAT'], sum(dict_tbhisdetiljual2[i]['JMLHRG_NETTO'])])
  
  list_pareto_nilai3 = []
  for i in dict_tbhisdetiljual3.keys():
    list_pareto_nilai3.append([i, dict_tbhisdetiljual3[i]['NAMA_OBAT'], sum(dict_tbhisdetiljual3[i]['JMLHRG_NETTO'])])
  
  