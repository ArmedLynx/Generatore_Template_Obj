import openpyxl
from openpyxl.utils import get_column_letter

from os import mkdir
from os.path import exists as path_exists

import sys, getopt


class Template():
    def __init__(self, template):
        self.template = template
        # Copio il file di testo in una stringa
        file = open(template, "rt")
        self.data = file.read()
        file.close()

    def Replace(self, tag, valore): # Sostituisce la stringa "tag" con la stringa "valore"
        self.data = self.data.replace(tag, valore)
    
    def Save(self, nome, out_dir="./"): # Salva il file come "nome" nel percorso "out_dir"
        outF = open(out_dir+"/"+nome, "w+")
        outF.write(self.data)
        outF.close()

class Data():
    def __init__(self, data):
        self.data = data
        # Apro il foglio di lavoro
        wb=openpyxl.load_workbook(self.data)
        self.sheet = wb.active
        # Ottengo una lista contenente le colonne utilizzate (Es.: ["A","B","C"])
        self.cols = [get_column_letter(c) for c in range(2, 1+len(tuple(self.sheet.columns)))]
    
    def GetFileName(self, riga): # Restituisce il filename contenuto nella cella alla riga "riga"
        return self.sheet["A"+str(riga)].value
    
    def GetFilenames(self): # Restitusce una stringa contenente tutti i filename ordinati
        names = [self.sheet["A"+str(i)].value for i in range(2, 1+len(tuple(self.sheet.rows)))]
        return names
    
    def GetTag(self, colonna): # Restituisce il tag contenuto nella cella alla colonna "colonna"
        return self.sheet[colonna+"1"].value
    
    def GetTags(self): # Restituisce una stringa contenente tutti i tag ordinati
        tags = [self.sheet[i+"1"].value for i in self.cols]
        return tags
    
    def GetValue(self, coord): # Restituisce il contenuto di una cella date le coordinate nel formato "A1"
        return self.sheet[coord].value
    
    def GetCell(self, file, tag): # Restituisce le coordinate della cella con colonna corrispondente a quella del tag e riga corrispondente a quella del filename
        col = self.cols[self.GetTags().index(tag)]
        row = self.GetFilenames().index(file)+2
        return col+str(row)
        

def main(argv):
    opts, args = getopt.getopt(argv, "t:d:o:Fh") # Eseguo il parsing della lista di argomenti per ottenere una lista di tuple (opzione, argomento)

    aiuto = '''
        -t <Path>\tPercorso per il file template
        -d <Path>\tPercorso per il file data
        -o <Path>\tPercorso per la cartella di output
        -F\tForza la riscrittura dei file di configurazione già creati
        -h\tAiuto
        '''
    # Assegno alle variabili il valore di default
    template_file = "./template.txt"
    data_file = "./data.xlsx"
    out_path = "./outputs"
    overwrite = False

    # Assegno alle variabili il valore passato da cmd
    for opt, arg in opts:
        if opt == "-t":
            template_file = arg
        elif opt == "-d":
            data_file = arg
        elif opt == "-o":
            out_path = arg
        elif opt == "-F":
            overwrite = True
        elif opt == "-h":
            print(aiuto)
            return
        else:
            print("Parametri non validi.\n"+aiuto)

    # Se la cartella di output non esiste la creo
    if path_exists(out_path) == False:
        mkdir(out_path)
        print(": Creo la cartella di output\n|")
    # Creo un oggetto contenente il folgio di lavoro xlsx
    xls = Data(data_file)
    print(": Leggo il file Excel\n|")
    # Eseguo un ciclo sulla lista contente i filename
    for file in xls.GetFilenames():
        # Se il flag overwrite è False e il file esiste non faccio nulla altrimenti creo il file configurazione
        if overwrite == False and path_exists(out_path+"/"+file+".txt") == True:
            print(": Il file "+file+".txt è già presente")
        else:
            # Creo un oggetto contenente il template
            conf = Template(template_file)
            print(":- Apro il template per "+file)
            # Eseguo un ciclo sulla lista contenete i tags
            for tag in xls.GetTags():
                # Sostituisco il tag con il contenuto della cella alle coordinate corrispondenti al file e al tag
                conf.Replace(tag, xls.GetValue(xls.GetCell(file, tag)))
                print(":-- Sostituisco "+tag+" con "+xls.GetValue(xls.GetCell(file, tag)))
            # Salvo il file di configurazione
            conf.Save(file+".txt", out_path)
            print(":- Salvo il file di cofigurazione "+file+"\n|")
    
    print(": Ho terminato la preparazione delle configurazioni")
    

if __name__ == "__main__":
    main(sys.argv[1:])