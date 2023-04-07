class Template():
    def __init__(self, template):
        self.template = template
        # Copio il file di testo in una stringa
        file = open(template, "rt")
        self.data = file.read()
        file.close()

    def Replace(self, tag, valore): # Sostituisce la stringa "tag" con la stringa "valore"
        self.data = self.data.replace(tag, str(valore))
    
    def Save(self, nome, out_dir="./"): # Salva il file come "nome" nel percorso "out_dir"
        outF = open(out_dir+"/"+nome, "w+")
        outF.write(self.data)
        outF.close()

    def Append(self, nome, out_dir="./"):
        outF = open(out_dir+"/"+nome, "a")
        self.data+="\n"
        outF.write(self.data)
        outF.close()