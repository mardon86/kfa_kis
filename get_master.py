import re,string

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
      master_list.append(line_list)
  return master_list

    
def write_master(master_list, dest_file):
  with open(dest_file, 'w') as df:
    for i in master_list:
      df.write(i[0])
      for j in i[1:len(i)]:
        df.write('\t' + j)
      df.write('\n')
      

if __name__ == '__main__':
  lrmf = load_raw_master('TBMASOBAT.TXT')
  patt = get_pattern(lrmf)
  lm = create_master_list(lrmf, patt)
  write_master(lm, 'MASTEROBAT.txt')