import tkinter as tk
from imports import *

# from imports import *

# options = ["Generate Supplier File", "Generate Address File", "Generate Supplier Site Data File"]
# variable = tk.StringVar(window)
# variable.set("Select File Type")

# option_menu = tk.OptionMenu(window, variable, *options)
# option_menu.grid(row=3, column=0, columnspan=3, sticky="n")


def bttn(window, row, column, text, bcolor, fcolor,bcolour_enter,  cmd, statee, photoo, heightt, widthh, font):
    # Adjust the desired height for the button
    mybutton = tk.Button(window, width=widthh, height=heightt,
                         fg=fcolor,
                         bg=bcolor,
                         border=0,
                         activeforeground=fcolor,
                         activebackground=bcolor,
                         command=cmd,
                         relief='solid',
                         borderwidth=0,
                         highlightthickness=1,
                         highlightbackground=bcolor,
                         state=statee,
                         compound='right',  # Set compound to 'left' for image on the left side of the text
                         image =photoo, text=text,
                         font=font)
    # mybutton.grid_propagate(False)  # Prevent grid from resizing the button
    # mybutton.configure(height=button_height)  # Set a fixed height for the button
    # mybutton.configure(height=10)  # Set a fixed width for the button
    # mybutton.configure(width=25)  # Set a fixed width for the button
    if mybutton['state'] == 'normal':
        mybutton['background'] = bcolour_enter
        mybutton['foreground'] = fcolor
    if mybutton['state'] == 'disabled':
        mybutton['background'] = bcolour_enter
        mybutton['foreground'] = fcolor

    def on_enter(e):
        if mybutton['state'] == 'normal':
            mybutton['background'] = bcolor
            mybutton['foreground'] = fcolor
        elif mybutton['state'] == 'disabled':
            mybutton['background'] = bcolour_enter
            mybutton['foreground'] = fcolor

    def on_leave(e):
        if mybutton['state'] == 'normal':
            mybutton['background'] = bcolour_enter
            mybutton['foreground'] = fcolor
        elif mybutton['state'] == 'disabled':
            mybutton['background'] = bcolour_enter
            mybutton['foreground'] = fcolor
    mybutton.bind("<Enter>", on_enter)
    mybutton.bind("<Leave>", on_leave)
    # mybutton.grid(row=row, column=column, rowspan=rowspan, columnspan=columnspan, pady=1, padx=1, sticky="nwes")
    mybutton.place(x=row, y=column)
    return mybutton

def opt_menu(window, row, column, rowspan, columnspan, text, options, bcolor, fcolor,bcolour_enter, cmd,statee,height,width):
    
    height = int(height/20) 
    myOptions = tk.OptionMenu(window, text, *options)
    myOptions.config( width=width, height=height, text=text,
                      fg=fcolor,
                      bg=bcolor,
                      border=0,
                      activeforeground=fcolor,
                      activebackground=bcolor,
                      command=cmd,
                      relief='solid',
                      borderwidth=1,
                      highlightthickness=0,
                      highlightbackground=bcolor,
                      state = statee)
    # print(myOptions['state'])
    if myOptions['state'] == 'normal':
            myOptions['background'] = bcolour_enter
            myOptions['foreground'] = fcolor
    if myOptions['state'] == 'disabled':
            myOptions['background'] = bcolour_enter
            myOptions['foreground'] = fcolor
    def on_enter(e):
        if myOptions['state'] == 'normal':
            myOptions['background'] = bcolour_enter
            myOptions['foreground'] = fcolor
        elif myOptions['state'] == 'disabled' :
            myOptions['background'] = bcolour_enter
            myOptions['foreground'] = fcolor
    def on_leave(e):
        if myOptions['state'] == 'normal':
            myOptions['background'] = bcolor
            myOptions['foreground'] = fcolor
        elif myOptions['state'] == 'disabled' :
            myOptions['background'] = bcolour_enter
            myOptions['foreground'] = fcolor
    myOptions.bind("<Enter>", on_enter)
    myOptions.bind("<Leave>", on_leave)
    myOptions.place( x = row, y = column)
    # myOptions.grid(row=row, column=column, rowspan=rowspan, columnspan=columnspan, pady=1, padx=1, sticky="nsew")
    return myOptions
# Create the main window
# window = Tk()
# window.geometry('300x500')
# window.configure(bg="#141414")

# Create the frame for the layout


# # Create the upper button
# bttn(frame, 0, 0, 2, 2, "Button 1", '#ffcc66', '#141414', None)

# # Create the lower buttons
# bttn(frame, 2, 0, 1, 1, "Button 2", '#ffcc66', '#141414', None)
# bttn(frame, 2, 1, 1, 1, "Button 3", '#ffcc66', '#141414', None)


# Run the main event loop
