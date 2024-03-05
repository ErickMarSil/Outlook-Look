import win32com.client                as pywin
import os 

class main :
    def __init__(self): # metodo init (inicializador) para instancias
        self.path = os.path.expanduser("~\\Documents\\") # espande as pastas para não ter erros de caminhos
        self.outlook = pywin.Dispatch("Outlook.Application").GetNamespace("MAPI") # cria objeto do tipo "Outlook"
        self.inbox = self.outlook.GetDefaultFolder(6)  # pega a primeira pasta "Caixa de entrada" do outlook
        self.messages = self.inbox.Items # pega todas as mensagens da caixa de entrada (self.inbox)
    
    def Take_FullP(self) -> str:
        return self.path + ""

    def start(self, subject ,exts, folders=[]) -> str: # metodo start começa a automação
        # retorna apenas o caminho completo do download para qualquer possivel automação
        self.path = self.Take_FullP()
        return self.instFile(self.lastFile(subject=subject), exts=exts, folders=folders)
    
    def lastFile(self, subject) -> str: # metodo lastFile procura o arquivo mais recente
        first:bool = True # verificador para loop posterior (para atribuir apenas o primeiro valor)
        list:vars = [] # lista de todos os emails que tenham o seu "assunto" igual ao (subject)
        # loop por todos os emails adicionando na lista (list) se tiver o assunto sujerido na variavel (subject)
        try: [list.append (email) for email in self.inbox.items 
              if subject in email.Subject and email.Unread == True]
        except: pass

        for email in list: # loop por emails para verificar o mais recente
            if first: # atribui primeiro valor para a variavel (bigger)
                bigger = email
                first = False 
            # verficações para mes e dia para novo (bigger)
            elif email.Senton.month >= bigger.Senton.month and  email.Senton.day > bigger.Senton.day:bigger = email
        return bigger

    def ThoughtFile(self, Mail, folders):
        FolderDestiny = self.outlook.Folders(1) # destino da pasta onde o arquivo será colocado (1) representa suas pastas
        if len(folders) > 0: # por padrão a pasta vem como nada caso não seja necessario move-lá
            # faz um loop criando o caminho para a pasta desejada e coloca o email no local
            for fold in folders: FolderDestiny = FolderDestiny.Folders[fold] 
            Mail.Move(FolderDestiny)

    def instFile(self, Content:vars , exts:str, folders:str=[]): # metodo instFile instala arquivo em Download e retorna o caminho completo de tal
        for attachment in Content.Attachments: # loop por todos os "anexos" (attachment) no corpo do email
            actExts = attachment.FileName[len(attachment.FileName)-4:len(attachment.FileName)] # pega extensão do arquivo
            if actExts == exts: # verifica se é o arquivo desejado
                Content.Unread = False # deixa como não lido
                attachment.SaveAsFile(os.path.join(self.path, str(attachment))) # instala arquivo
                self.ThoughtFile(Mail=Content, folders=folders) # joga email para pasta necessaria
                return os.path.join(self.path, str(attachment)) # retorna caminho completo
            
SaverInstance = main() # instancia do robo
subject = "" # coloque oque normalmente aparece como titulo do email

def sentry(subject="") -> bool:
    inbox = pywin.Dispatch("Outlook.Application").GetNamespace("MAPI").GetDefaultFolder(6)
    return True in [subject in mail.Subject and mail.Unread == True for mail in inbox.Items]

def start(pathFile):
    pathDoc = os.path.expanduser("~\\Documents\\") # coaminho do arquivo texto para colocar o camindo do download
    with open(pathDoc + "pathFile.txt", "w") as file: file.truncate(0), file.write(pathFile), file.close # salva arquivo
# verifica se tem um email novo ou não
if sentry(subject) == True: start(SaverInstance.start(subject , "xlsb", [""]))