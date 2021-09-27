import tkinter as tk
from PIL import ImageTk, Image
from file_transform import transformer

def main_window():
    # window creation and widgets
    window = tk.Tk()
    window.title('Bible Groupe Aline')

    # Centers the window
    w = window.winfo_reqwidth()
    h = window.winfo_reqheight()
    positionRight = int(window.winfo_screenwidth() / 2 - w / 2)
    positionDown = int(window.winfo_screenheight() / 2 - h / 2)
    window.geometry("710x500".format(positionRight, positionDown))

    # window is not resizable
    window.resizable(width=False, height=False)

    # Create a photo-image object of the image in the path
    image1 = Image.open("Images/unicolor_blue.jpg")
    image2 = Image.open("Images/Groupe aline.png")
    image1 = image1.resize((150, 1000), Image.ANTIALIAS)  # resize image
    image2 = image2.resize((100, 100), Image.ANTIALIAS)  # resize image
    data_image1 = ImageTk.PhotoImage(image1)
    data_image2 = ImageTk.PhotoImage(image2)
    label1 = tk.Label(image=data_image1)
    label2 = tk.Label(image=data_image2)
    label1.image = data_image1
    label2.image = data_image2

    # Position image
    label1.place(x=0, y=0)
    label2.place(x=20, y=10)

    # Create labels for title and intro
    title = tk.Label(text="Page 1 : Bible Groupe Aline", font=('Helvetica', 15, 'bold'))
    text_intro = tk.Label(text="Bienvenue dans l'outil de compilation de la bible du groupe Aline.\n"
                               " Veuillez cliquer sur les boutons et choisir les fichiers à traiter",font=('Helvetica', 10, 'italic'))

    # Create buttons to choose files and call setFile function
    Analcli1 = tk.Button(text="Analcli annuel de l'année en cours", font=('Helvetica', 10),command=lambda: transformer.setFile(file1_label, 0))
    Analcli2 = tk.Button(text="Analcli annuel de l'année suivante", font=('Helvetica', 10),command=lambda: transformer.setFile(file2_label, 1))
    Analcli3 = tk.Button(text="Analcli annuel de l'année en cours", font=('Helvetica', 10),command=lambda: transformer.setFile(file3_label, 2))
    Analcli4 = tk.Button(text="Analcli mensuel de l'année suivante", font=('Helvetica', 10),command=lambda: transformer.setFile(file4_label, 3))
    button_suivant = tk.Button(text="Suivant", font=('Helvetica', 10, 'italic'), bg="#75D5E5",command=(lambda: second_window(window)))

    # Create labels for the chosen files
    file1_label = tk.Label(text="-", font=('Helvetica', 10))
    file2_label = tk.Label(text="-", font=('Helvetica', 10))
    file3_label = tk.Label(text="-", font=('Helvetica', 10))
    file4_label = tk.Label(text="-", font=('Helvetica', 10))

    # Placing of widgets
    title.place(x=310, y=20)
    text_intro.place(x=250, y=60)

    # Button placement
    Analcli1.place(x=290, y=130, width=300)
    Analcli2.place(x=290, y=210, width=300)
    Analcli3.place(x=290, y=290, width=300)
    Analcli4.place(x=290, y=370, width=300)
    button_suivant.place(x=520, y=440, width=160)

    # The label placement
    file1_label.place(x=250, y=160)
    file2_label.place(x=250, y=240)
    file3_label.place(x=250, y=320)
    file4_label.place(x=250, y=400)

    # main loop
    window.mainloop()
    return window

def second_window(window):
    window.destroy()

    second_window = tk.Tk()
    second_window.title('Bible Groupe Aline')

    # Centers the window
    w = second_window.winfo_reqwidth()
    h = second_window.winfo_reqheight()
    positionRight = int(second_window.winfo_screenwidth() / 2 - w / 2)
    positionDown = int(second_window.winfo_screenheight() / 2 - h / 2)
    second_window.geometry("710x500".format(positionRight, positionDown))

    # window is not resizable
    second_window.resizable(width=False, height=False)

    # Create a photo-image object of the image in the path
    image1 = Image.open("Images/unicolor_blue.jpg")
    image2 = Image.open("Images/Groupe aline.png")
    image1 = image1.resize((150, 1000), Image.ANTIALIAS)  # resize image
    image2 = image2.resize((100, 100), Image.ANTIALIAS)  # resize image
    data_image1 = ImageTk.PhotoImage(image1)
    data_image2 = ImageTk.PhotoImage(image2)
    label1 = tk.Label(image=data_image1)
    label2 = tk.Label(image=data_image2)
    label1.image = data_image1
    label2.image = data_image2

    # Position image
    label1.place(x=0, y=0)
    label2.place(x=20, y=10)

    # Create labels for title and intro
    title = tk.Label(text="Page 2 : Bible Groupe Aline ", font=('Helvetica', 15, 'bold'))
    text_intro = tk.Label(text="Bienvenue dans l'outil de compilation de la bible du groupe Aline.\n"
                               " Veuillez cliquer sur les boutons et choisir les fichiers à traiter",font=('Helvetica', 10, 'italic'))

    # Placing of widgets
    title.place(x=310, y=20)
    text_intro.place(x=250, y=60)

    # window is not resizable
    second_window.resizable(width=False, height=False)

    Rep_Client_button = tk.Button(text="Rep Client", font=('Helvetica', 10),command=lambda: transformer.setFile(file5_label, 4))
    Categorie_Client_button = tk.Button(text="Catégorie client", font=('Helvetica', 10),command=lambda: transformer.setFile(file6_label, 5))
    PPN_button = tk.Button(text="PPN Client", font=('Helvetica', 10),command=lambda: transformer.setFile(file7_label, 6))
    Class_produit_button = tk.Button(text="Classification produit", font=('Helvetica', 10),command=lambda: transformer.setFile(file8_label, 7))
    button_compile = tk.Button(text="Suivant", font=('Helvetica', 10, 'italic'), bg="#75D5E5", command=(lambda:transformer.convert_files('self')))

    Rep_Client_button.place(x=290, y=130, width=300)
    Categorie_Client_button.place(x=290, y=210, width=300)
    PPN_button.place(x=290, y=290, width=300)
    Class_produit_button.place(x=290, y=370, width=300)
    button_compile.place(x=520, y=440, width=160)

    file5_label = tk.Label(text="-", font=('Helvetica', 10))
    file6_label = tk.Label(text="-", font=('Helvetica', 10))
    file7_label = tk.Label(text="-", font=('Helvetica', 10))
    file8_label = tk.Label(text="-", font=('Helvetica', 10))

    file5_label.place(x=250, y=160)
    file6_label.place(x=250, y=240)
    file7_label.place(x=250, y=320)
    file8_label.place(x=250, y=400)

    # main loop
    second_window.mainloop()
