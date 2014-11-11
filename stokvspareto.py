import re, string

def take_stock_full(raw_stock_file):
  with open(raw_stock_file,'r') as f:
    stockfull = []
    for line in f:
      if len(line) > 5:
        if line[5].isdigit():
          stockfull.append([re.sub(' *$','',string.split(line,'\xb3')[2]), re.sub('^ *| *$','',string.split(line,'\xb3')[3]), re.sub('^ *| *$','',string.split(line,'\xb3')[4]), re.sub('^ *|,|\.\d+','',string.split(line,'\xb3')[5]), re.sub('^ *|,|\.\d+','',string.split(line,'\xb3')[6])])
  return stockfull
  
def take_pareto(raw_pareto_file):
  with open(raw_pareto_file,'r') as f:
    pareto_list = []
    for line in f:
      if len(line) > 4:
        if line[4].isdigit():
          a = re.sub(' *$','',string.split(line,'\xb3')[2])          # nama barang
          b = re.sub(' *$','',string.split(line,'\xb3')[4])          # nama pabrik
          c = re.sub('^ *|,|\.\d+','',string.split(line,'\xb3')[6])  # qty
          d = re.sub('^ *|,|\.\d+','',string.split(line,'\xb3')[7])  # nilai
          e = re.sub('^ *','',string.split(string.split(line,'\xb3')[10],'-')[0])       # freq
          pareto_list.append([a,b,c,d,e])
    nama_barang = []
    for i in pareto_list:
      nama_barang.append(i[0])
    nama_barang.sort()
    curs = 0
    mult_elements = []
    while curs < len(nama_barang) - 1:
      if nama_barang[curs] == nama_barang[curs + 1]:
        mult_elements.append(nama_barang[curs])
      curs += 1
    mult_elements_set = set(mult_elements)
    mult_elements = list(mult_elements_set)
    mult_elements_index = {}
    for i in mult_elements:
      mult_elements_index[i] = [ j for j,k in enumerate(pareto_list) if k[0] == i ]
    for i in mult_elements:
      for j in mult_elements_index[i]:
        if pareto_list[j][1] == "NAMA_PABRIK tidak ada":
          del pareto_list[j]
    pareto = []
    for i in pareto_list:
      pareto.append([i[0], i[2], i[3], i[4]])
    return pareto

def obtain_procurement_plan(stock, pareto, dest_file):
  stock_dict = {}
  pareto_dict = {}
  rencana_pengadaan = {}
  with open(dest_file,'w') as rp:
    for i in stock:
      stock_dict[i[0]] = [i[1], i[2], i[3], i[4]]
    for i in pareto:
      pareto_dict[i[0]] = [i[1], i[2], i[3]]
    for i in stock_dict.keys():
      if i in pareto_dict.keys():
        rencana_pengadaan[i] = [stock_dict[i][0],stock_dict[i][1],stock_dict[i][2],stock_dict[i][3],pareto_dict[i][0],pareto_dict[i][1],pareto_dict[i][2]]
      else:
        rencana_pengadaan[i] = [stock_dict[i][0],stock_dict[i][1],stock_dict[i][2],stock_dict[i][3],'0','0','0']
    for i in pareto_dict.keys():
      if i not in stock_dict.keys():
        rencana_pengadaan[i] = ['','','0','',pareto_dict[i][0],pareto_dict[i][1],pareto_dict[i][2]]
    rp.write('NO.\tNAMA BARANG\tLOKASI\tSATUAN\tSTOK\tHNA\tJUMLAH\tNILAI JUAL\tKUNJUNGAN\n')
    counting = 1
    for i in rencana_pengadaan.keys():
      if (rencana_pengadaan[i][2] != '0' and rencana_pengadaan[i][2] != '') or (rencana_pengadaan[i][4] != '0' and rencana_pengadaan[i][4] != ''):
        rp.write(str(counting) + '\t' + i + '\t' + rencana_pengadaan[i][0] + '\t' + rencana_pengadaan[i][1] + '\t' + rencana_pengadaan[i][2] + '\t' + rencana_pengadaan[i][3] + '\t' + rencana_pengadaan[i][4] + '\t' + rencana_pengadaan[i][5] + '\t' + rencana_pengadaan[i][6] + '\n')
        counting += 1
      
if __name__ == '__main__':
  stock = take_stock_full('shaft.txt')
  pareto = take_pareto('pareto.txt')
  obtain_procurement_plan(stock, pareto, 'stok_vs_pareto.txt')
      
