import re, string
from operator import itemgetter

#Script ini digunakan untuk membagi barang yang masuk ke pareto A, B, dan C


def take_pareto(raw_pareto_file):
  with open(raw_pareto_file,'r') as f:
    pareto_dict = {}
    for line in f:
      if len(line) > 4:
        if line[4].isdigit():
          a = re.sub(' *$','',string.split(line,'\xb3')[2])          # nama barang
          b = re.sub('^ *|,|\.\d+','',string.split(line,'\xb3')[6])  # qty
          c = re.sub('^ *|,|\.\d+','',string.split(line,'\xb3')[7])  # nilai
          d = re.sub('^ *','',string.split(string.split(line,'\xb3')[10],'-')[0])       # freq
          pareto_dict[a] = [int(b),int(c),int(d)]
  return pareto_dict

def get_all_data_pareto():
  semesta = []
  par_123 = {}
  for i in range(1,4):
    par_123[i] = take_pareto("pareto{0}.txt".format(i))
  for i in range(1,4):
    for j in par_123[i].keys():
      if j not in semesta:
        semesta.append(j)
  return (semesta,par_123)

def get_sorted_100(the_list, itmgttr):
  return sorted(the_list, key=itemgetter(itmgttr), reverse = True)

def get_nilai_80(the_list):
  tl = the_list
  sum_nilai = 0
  for i,j,k,l in tl:
    sum_nilai += k
  nilai_80 = []
  sum_nilai_80_percent = 0.8 * sum_nilai
  sum_until80 = 0
  for i in tl:
    if sum_until80 < sum_nilai_80_percent:
      nilai_80.append(i)
      sum_until80 += i[2]
    else:
      break
  return nilai_80

def get_kunj_80(the_list):
  tl = the_list
  sum_kunj = 0
  for i,j,k,l in tl:
    sum_kunj += l
  kunj_80 = []
  sum_kunj_80_percent = 0.8 * sum_kunj
  sum_until80 = 0
  for i in tl:
    if sum_until80 < sum_kunj_80_percent:
      kunj_80.append(i)
      sum_until80 += i[3]
    else:
      break
  return kunj_80

def get_pareto_A(nil80, kunj80):
  pareto_A_items = []
  
  nil80_items = []
  for i in nil80:
    nil80_items.append(i[0])
  
  kunj80_items = []
  for i in kunj80:
    kunj80_items.append(i[0])
  
  for i in nil80_items:
    if (i in kunj80_items):
      pareto_A_items.append(i)
  
  return pareto_A_items
  

class Items:
  def __init__(self, nama_obat, data, data_pareto):
    self.nama_obat = nama_obat
    self.data = data
    self.data_pareto = data_pareto
    self.qty_penj = []
    self.kunj = []
    self.pareto = get_pareto()
    self.get_data()
  def get_data(self):
    for i in range(1,4):
      if nama_obat in self.data[i].keys():
        self.qty_penj.append(self.data[i][nama_obat][0])
        self.kunj.append(self.data[i][nama_obat][2])
      else:
        self.qty_penj.append(0)
        self.kunj.append(0)
  def get_max_qty(self):
    return max(self.qty_penj)
  def get_cure_qty(self):
    return sum(self.qty_penj)/sum(self.kunj)
  def get_pareto(self):
    if nama_obat in self.data_pareto[0]:
      return "A"
    elif nama_obat in self.data_pareto[1]:
      return "B"
    else:
      return "C"
      


if __name__ == "__main__":
  S_items, par_123 = get_all_data_pareto()
  
  listitems_123 = {}
  for i in range(1,4):
    listitems_123[i] = []
    for j in par_123[i].keys():
      listitems_123[i].append([j, par_123[i][0], par_123[i][1], par_123[i][2]])
  
  nilai_100_123 = {}
  for i in range(1,4):
    nilai_100_123[i] = get_sorted_100(listitems_123[i], 2)
  
  kunj_100_123 = {}
  for i in range(1,4):
    kunj_100_123[i] = get_sorted_100(listitem_123[i], 3)

  nilai_80_123 = {}
  for i in range(1,4):
    nilai_80_123[i] = get_nilai_80(nilai_100_123[i])
    
  kunj_80_123 = {}
  for i in range(1,4):
    kunj_80_123[i] = get_kunj_80(kunj_100_123[i])
  
  pareto_A_items_123 = {}
  for i in range(1,4):
    pareto_A_items_123[i] = get_pareto_A(nilai_80_123[i], kunj_80_123[i])
  
  pareto_A_items_unity = []
  for i in range(1,4):
    for j in pareto_A_items_123[i]:
      if j not in pareto_A_items_unity:
        pareto_A_items_unity.append(j)
  
  nil80_U_kunj80 = []
  for i in range(1,4):
    for j in nilai_80_123[i]:
      if j[0] not in nil80_U_kunj80:
        nil80_U_kunj80.append(j[0])
        
    for k in kunj_80_123[i]:
      if k[0] not in nil80_U_kunj80:
        nil80_U_kunj80.append(k[0])
  
  pareto_B_items_unity = []
  for i in nil80_U_kunj80:
    if i not in pareto_A_items_unity:
      pareto_B_items_unity.append(i)
  
  pareto_C_items_unity = []
  for i in S_items:
    if i not in nil80_U_kunj80:
      pareto_C_items_unity.append(i)
  
  data_pareto_items = (pareto_A_items_unity, pareto_B_items_unity, pareto_C_items_unity)

  with open('pareto_a.txt','w') as pa:
    for i in pareto_A_items_unity:
      pa.write(i + "\n")

  with open('pareto_b.txt','w') as pb:
    for i in pareto_B_items_unity:
      pb.write(i + "\n")

  with open('pareto_c.txt','w') as pc:
    for i in pareto_C_items_unity:
      pc.write(i + "\n")
  
  with open('semesta.txt','w') as s:
    for i in S_items:
      s.write(i + "\n")