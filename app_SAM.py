from xlwt import *
from xlrd import open_workbook
from tkinter import *
from tkinter import filedialog
# import httplib2
import os
import xml.etree.ElementTree as etree
import re

url_SAM_API = "http://139.54.78.41:8080/xmlapi/invoke"
url_xls_output ="C:\\Users\\afourmy\\Desktop\\SAM-O\\Responses-xls-formatting\\Nodes\\find_processing.xls"

# variable globales: 
path_xml_input = "C:\\Users\\afourmy\\Desktop\\SAM-O\\SAM-2 requests\\Network Elements\\Find.xml"
path_xml_output = "C:\\Users\\afourmy\\Desktop\\SAM-O\\SAM-2 requests\\Network Elements\\output.xml"
path_xls = "C:\\Users\\afourmy\\Desktop\\SAM-O\\SAM-2 requests\\Network Elements\\findNE.xls"
SAM_address = "http://139.54.78.41:8080/xmlapi/invoke"
liste_classes = ["netw.RadioPhysicalLink", "netw.NetworkElement", "netw.DiscoveredPhysicalLink","equipment.PhysicalPort", "equipment.LogicalPort", "netw.PhysicalInterfaceCtp", "bundle.Interface", "netw.Connection", "netw.LogicalInterface", "netw.PhysicalLink", "netw.StatefullConnectableInterface", "netw.RouterTerminatingIpInterface", "equipment.CardSlot", "equipment.Slot", "lag.Interface", "netw.ConnectableInterface", "vprn.L3AccessInterface", "lte.ENBEquipment", "lte.EPSPath", "lteggsn.SgwGaPath", "lte.S1uPath", "lte.S1mmePath", "radioequipment.RadioPortSpecifics", "ethring.RadioRing"]

default_port = 8080


def HTTP_post_request(contenu1, contenu2, information):
    pass
    # httplib2.debuglevel = 1
    # http = httplib2.Http()
    # headers = {}
    # # HTTP This port provides an HTTP interface for 5620 SAM-O clients to access the 5620 SAM server.
    # # Port 8443: HTTPS This port provides an HTTPS (secure HTTP) interface for 5620 SAM-O clients that wish to use this protocol to access the 5620 SAM server 
    # 
    # # Change the login / password in the input xml file
    # input_tree = etree.parse(path_xml_input)
    # 
    # 
    # request = open(path_xml_input,"rb").read()

   ##   
    # # The request() method returns two values. The first is an httplib2.Response object, which contains all the HTTP headers the server returned. For example, a status code of 200 indicates that the request was successful.
    # # The content variable contains the actual data that was returned by the HTTP server. The data is returned as a bytes object, not a string. Character encoding to be determined for a conversion to string
    # try: 
    #     response, content = http.request(SAM_address, 'POST', body=request)
    #     # on voit dans response (headers) que l'encodage est charset=ISO-8859-1
    #     ecriture = open(path_xml_output, "w")
    #     ecriture.write(str(content)[2:-1])
    #     ecriture.close()
    # 
    #     contenu1.delete("1.0",END) # je vire le contenu précédent, pour que les requêtes ne s'accumulent pas
    #     contenu2.delete("1.0",END)
    #     contenu1.insert(END, content)
    #     contenu2.insert(END, response)
    #     
    #     information.configure(text = "")
    # 
    # except ConnectionRefusedError: 
    #     contenu1.delete("1.0",END)
    #     contenu1.insert(END, "ConnectionRefusedError: [WinError 10061] No connection could be made because the target machine actively refused it\n" + "Le SAM refuse parfois l'envoi de requêtes sur le port 8443")
    #     
    # except OSError: 
    #     contenu1.delete("1.0",END)
    #     contenu1.insert(END, "OSError: [WinError 10049] The requested address is not valid in its context\n" + "Introduisez une adresse IP valide dans File > Default parameters > SAM IP Address")
    #     
    # except httplib2.ServerNotFoundError:
    #     contenu1.delete("1.0",END)
    #     contenu1.insert(END, "socket.gaierror: [Errno 11004] getaddrinfo failed: unable to find the server\n" + "Introduisez une adresse IP valide dans File > Default parameters > SAM IP Address") 

    
# définition de l'action à effectuer si l'utilisateur clique sur update: on écrit les valeurs du SAM dans le fichier excel :
def update_excel(list_entree,information):
    global liste_classes
    
    tree = etree.parse(path_xml_output)
    # création
    book = Workbook()
    
    # création de la feuille 1
    # to get rid of overwrite forbidden exception: worksheet = workbook.add_sheet("Sheet 1", cell_overwrite_ok=True)
    nom_scenario = list_entree[0].get()
    feuil1 = book.add_sheet(nom_scenario, cell_overwrite_ok=True)
    list_col = [feuil1.col(k) for k in range(50)]
    for col in list_col:
        col.width = 256 * 25 # 25 characters wide
    
    k=9 # on commence à -9 car les premiers tags jusqu'à "result" ne servent à rien ici. Cela évite de laisser de l'espace blanc dans la feuille excel à cause de l'itération k+=1. Cela étant, Spider permet de spécifier la ligne à laquelle commence les données dans Generic I/E, donc ce n'était pas vraiment un problème. 
    list_element = [entree.get() for entree in list_entree]
    for p in range(1,32):
        feuil1.write(0,p,list_element[p])
    
    feuil = feuil1
    valeur = ""
    for node in tree.iter():
        tag = node.tag[12:] # c'est ce qui se trouve entre "< >". Commmence par "{xmlapi_1.0}" donc on retire les 12 premiers caractères
        if(tag not in liste_classes):
            for q in range(0,len(list_entree)):
                if(list_element[q] == tag):
                    if(tag == "pointer"):
                        valeur = valeur + str(node.text)
                        feuil.write(k,q,valeur) 
                    else:
                        valeur = ""
                        feuil.write(k,q,node.text)
        else: 
            if(k<30000):
                k+=1
            else:
                nom_scenario = nom_scenario+"1"
                feuil = book.add_sheet(nom_scenario, cell_overwrite_ok=True)
                k=9
                
    
    feuil1.write(0,0,"nom_scenario") # pour la première colonne
    for x in range(1,k+1):
        feuil1.write(x,0,nom_scenario)
    # création matérielle du fichier résultant
    book.save(path_xls)
    # methode configure du widget chaine pour modifier son attribut "text"
    maj = path_xls + " updated"
    information.configure(text = maj)
    
def default_path():
    # variables globales
    global path_xml_input
    global path_xml_output
    global path_xls
    global SAM_address
    global default_port
    
    newfenetre = Toplevel()
    newfenetre.geometry("600x245")
    newfenetre.title("Default paths used by the application")
    #fenetre.wm_attributes("-topmost",1) # keep the window always on top (finalement j'utilise focus_force() dans file_path, sinon l'explorateur est en dessous de la default paths window...

    # affichage des paths définis par l'utilisateur. Affichage par défaut: valeur des global variables associées.
    Var_xml_input = StringVar()
    Var_xml_input.set(path_xml_input)
    entry_path_xml_input = Entry(newfenetre, textvariable = Var_xml_input, width=80)
    
    Var_xml_output = StringVar()
    Var_xml_output.set(path_xml_output)
    entry_path_xml_output = Entry(newfenetre, textvariable = Var_xml_output, width=80)
    
    Var_xls = StringVar()
    Var_xls.set(path_xls)
    entry_path_xls = Entry(newfenetre, textvariable = Var_xls, width=80)
    label_SAM = Label(newfenetre, text="SAM IP address :")
    label_XML = Label(newfenetre, text="XML Java Class :")
    
    # addresse du SAM (définie par l'utilisateur, 255.255.255.255 by default 
    SAM_var = StringVar()
    SAM_var.set("139.54.60.40")
    SAM_IP = Entry(newfenetre, textvariable = SAM_var, width=30)
    
    # Java class used in the xml query
    Javaclass_var = StringVar()
    Javaclass_var.set("netw.NetworkElement")
    Javaclass = Entry(newfenetre, textvariable = Javaclass_var, width=30)
    
    # selection des paths par l'utilisateur
    bouton_xml_input = Button(newfenetre, text='XML Request', command = lambda: file_path(newfenetre,entry_path_xml_input), width=12, height=1)
    bouton_xml_output = Button(newfenetre, text='XML Response', command = lambda: file_path(newfenetre,entry_path_xml_output), width=12, height=1)
    bouton_xls = Button(newfenetre, text='XLS Export', command = lambda: file_path(newfenetre,entry_path_xls), width=12, height=1)
    bouton_save_default_paths = Button(newfenetre, text='Save', command = lambda: save_paths(newfenetre,entry_path_xml_input,entry_path_xml_output,entry_path_xls,SAM_IP,label_verdict,valeur_port,Javaclass), width=12, height=1)
    bouton_close = Button(newfenetre, text="Close", command = newfenetre.destroy, width=12, height=1)
    
    # affichage des boutons / label dans la grille
    bouton_xml_input.grid(row=0,column=0, pady=5, padx=5, sticky=W)
    bouton_xml_output.grid(row=1,column=0, pady=5, padx=5, sticky=W)
    bouton_xls.grid(row=2,column=0, pady=5, padx=5, sticky=W)
    entry_path_xml_input.grid(row=0, column=1, sticky=W)
    entry_path_xml_output.grid(row=1, column=1, sticky=W)
    entry_path_xls.grid(row=2, column=1, sticky=W)
    label_SAM.grid(row=3,column=0, pady=5, padx=5, sticky=W)
    SAM_IP.grid(row=3, column=1, pady=5, padx=5, sticky=W)
    label_XML.grid(row=4, column=0, pady=5, padx=5, sticky=W)
    Javaclass.grid(row=4, column=1, pady=5, padx=5, sticky=W)
    
    # radio button pour choisir si on envoie la requête au SAM sur le port 8080 ou 8443
    valeur_port = IntVar()
    valeur_port.set(default_port) 
    port_standard = Radiobutton(newfenetre, text="Port 8080", variable=valeur_port, value=8080, command=lambda: valeur_port)
    port_https = Radiobutton(newfenetre, text="Port 8443", variable=valeur_port, value=8443, command=lambda: valeur_port)
    
    # affichage des radio button
    port_standard.grid(row=5,column=0, pady=5, padx=5, sticky=W)
    port_https.grid(row=5,column=1, pady=5, padx=5, sticky=W)
    
    # bouton save et close sur la dernière ligne
    bouton_close.grid(row=6,column=1, pady=5, padx=5, sticky=E)
    bouton_save_default_paths.grid(row=6,column=0, sticky=W, pady=5, padx=5)
    
    # label pour indiquer que la sauvegarde a réussi
    label_verdict = Label(newfenetre, text="")
    label_verdict.grid(row=6, column=1, pady=5, padx=5, sticky=W)
    
def save_paths(newfenetre,entry_path_xml_input,entry_path_xml_output,entry_path_xls,SAM_IP,label_verdict,port,javaclass):
    # variables globales
    global path_xml_input 
    global path_xml_output
    global path_xls
    global SAM_address
    global default_port
    global liste_classes
    
    path_xml_input = entry_path_xml_input.get()
    path_xml_output = entry_path_xml_output.get()
    path_xls = entry_path_xls.get()
    default_port = port.get()
    IP = SAM_IP.get()
    
    SAM_address = "http://%s:%s/xmlapi/invoke" % (IP, default_port)
    liste_classes.append(javaclass.get())
    
    label_verdict.configure(text = "Nouveaux paramètres enregistrés avec succès")

def set_default_parameters(liste_default, liste_entree):
    for k in range(len(liste_default)):
        liste_entree[k+1].delete(0, END)
        liste_entree[k+1].insert(0,liste_default[k])

def init_app():
    # ----- Programme principal : -----
    taille = 39
    # Tk est l'une des classes du module tkinter. en faisant Tk(), on en crée une instance.
    fenetre = Tk()
    fenetre.title("Conversion 5620 SAM-O -> Spider")
    
    menubar = Menu(fenetre)
    filemenu = Menu(menubar, tearoff=0)
    filemenu.add_command(label="Memo", command=lambda: memo())
    filemenu.add_command(label="Reset fields", command=lambda: reset_fields(liste_entree))
    filemenu.add_command(label="Default parameters", command=lambda: default_path())
    filemenu.add_separator()
    filemenu.add_command(label="Exit", command=fenetre.destroy)
    menubar.add_cascade(label="File",menu=filemenu)
    filemenu2 = Menu(menubar, tearoff=0)
    filemenu2.add_command(label="Write to xls", command=lambda: update_excel(liste_entree,information))
    filemenu2.add_command(label="Send request", command=lambda: HTTP_post_request(contenu1, contenu2, information))
    menubar.add_cascade(label="Actions",menu=filemenu2)
    
    fenetre.config(menu=menubar)
    # on crée deux textfields pour que l'utilisateur puisse entrer les noms des noeuds du fichier xml qu'il souhaite avoir dans le fichier excel
    
    champ_scenario = Label(fenetre, text = "Nom du scenario")
    list_champ = [champ_scenario]
    for champ_index in range(taille):
        label = Label(fenetre, text="Parametre %s" % (champ_index+1))
        list_champ.append(label)
    

    
    # # commande exécutée par le programme lorsque l'utilisateur actionnera la touche Return (ou Enter): quand l'utilisateur presse "entrer", la fonction évaluer s'exécute
    # # si on fait appel à la fonction display element en utilisant bind, il faut IMPERATIVEMENT qu'elle ait en argument "event"
    # #entree.bind("<Return>", display_element)
    
    # on pourrait définir un texte dans label, mais on ne le fait que dans la fonction display element pour préciser que le fichier excel est mis à jour
    information = Label(fenetre, text="")
    
    # le widget text permet autowraps le texte (il n'y a pas vraiment d'autowrapping avec les autres widgets)
    contenu1 = Text(fenetre, height=10, width=131, wrap=WORD)
    contenu2 = Text(fenetre, height=5, width=131, wrap=WORD)
    

    # la méthode pack() réduit automatiquement la taille de la fenêtre « maître » afin qu'elle soit juste assez grande pour contenir les widgets « esclaves » définis au préalable.
    # l'ordre dans lequel on les met est l'ordre dans lequel les widgets apparaissent dans la fenêtre. Problème: c'est très peu flexible. On utilise "grid" pour disposer les objets comme dans un tableau.
    


    information.grid(columnspan = 7)
    contenu1.grid(columnspan = 12)
    contenu2.grid(columnspan = 12)

    
    # mainloop: provoque le démarrage du réceptionnaire d'événements associé à la fenêtre.
    fenetre.mainloop()

def main():
    init_app()

if __name__ == "__main__": main()