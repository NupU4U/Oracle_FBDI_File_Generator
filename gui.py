from imports import *
from button_front_page import *
import customtkinter

def change_state(mybutton):
    if mybutton['state'] == 'disabled':
        mybutton.configure(state='normal')
    else:
        mybutton.configure(state='disabled')
        
def save_dict_as_md(dictionary):
    root = Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(filetypes=[('Markdown Files', '*.md')])
    if file_path:
        if not file_path.lower().endswith('.md'):
            file_path += '.md'
        with open(file_path, 'w') as file:
            for key, value in dictionary.items():
                file.write(f"## {key}\n\n")
                file.write(f"{value}\n\n")
        # print("Dictionary saved successfully.")
        messagebox.showinfo("EY DATA CONVERSION TOOL", "Mappings saved!\n\nPlease press Next")

    else:
        messagebox.showinfo("EY DATA CONVERSION TOOL", "Aborted by User!\n\nPlease press Next")
        # print("File save canceled.")

def load_dict_from_md():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[('Markdown Files', '*.md')])
    if file_path:
        dictionary = {}
        with open(file_path, 'r') as file:
            lines = file.readlines()
            key = None
            value = ''
            for line in lines:
                line = line.strip()
                if line.startswith('##'):
                    if key is not None:
                        dictionary[key] = value.strip()
                    key = line[2:].strip()
                    value = ''
                else:
                    value += line + '\n'
            if key is not None:
                dictionary[key] = value.strip()
        # print("Dictionary loaded successfully.")
        return dictionary
    else:
        messagebox.showinfo("EY DATA CONVERSION TOOL", "Aborted by User!\n\nPlease press Next")
        # print("File open canceled.")
        return None

def upload_excel():
    global data_list
    global supplier_consolidated
    # Open file dialog to select the Excel file
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        # Load the workbook
        workbook = openpyxl.load_workbook(file_path)

        # Select the active sheet
        sheet = workbook.active

        # Retrieve the column headings
        supplier_consolidated = [cell.value for cell in sheet[1]]

        # Store the data column-wise
        column_data = []
        for column in sheet.iter_cols(min_row=2, values_only=True):
            column_data.append(list(column))

        # Create a list of tuples with column headings and data
        data_list = list(zip(supplier_consolidated, column_data))
        # return data_list
        # Print the data
        
        # btn_next_1.forget()
        btn_next.place(x=100,y=270)
        # option_menu_file.place_forget()
        # option_menu_supplier.place(x=155,y=155)
        change_state(btn_upload)
        change_state(option_menu_file)
        messagebox.showinfo("EY DATA CONVERSION TOOL", "Upload Successful!\n\nPlease press Next")
        # for column_heading, column_values in data_list:
        #     print(f'{column_heading}: {column_values}')
        # print(supplier_consolidated)
        # print(data_list)
def show_type_of_files():
    global selected_file_type 
    btn_next.place_forget()
    btn_upload.place_forget()
    option_menu_file.place_forget()
    if selected_file_type == "Supplier":
        option_menu_supplier.place(x=25,y=150)
    elif selected_file_type == "Project":
        option_menu_project.place(x=25,y=150)
    btn_next1.place(x = 25, y =202 )
    btn_back1.place(x = 100 , y = 270)
    def enable_button2(*args):
        global selected_file
        if var_file.get():
            btn_next1.config(state=tk.NORMAL)
        else:
            btn_next1.config(state=tk.DISABLED)
        selected_file = var_file.get() 
        # print(selected_file)   
    var_file.trace("w", enable_button2)
        
def back1():
    global var_type_of_file,  mappings,mappings_supplier,mappings_supplier_address,mappings_supplier_site
    mappings_supplier.clear()
    mappings.clear()
    mappings_supplier_address.clear()
    mappings_supplier_site.clear()
    option_menu_supplier.place_forget()
    btn_next1.place_forget()
    btn_back1.place_forget()
    supplier_consolidated.clear()
    data_list.clear()
    btn_upload.place(x= 25, y = 202)
    option_menu_file.place(x= 25, y = 150)
    var_type_of_file.set("Select File Type")
    change_state(option_menu_file)
    change_state(btn_upload)
    def enable_button1(*args):
    # Check if an option is selected
        if var_type_of_file.get():
            btn_upload.config(state=tk.NORMAL)
        else:
            btn_upload.config(state=tk.DISABLED)

        # print( selected_file)
        # print(mappings)
    var_file.trace("w", enable_button1)

def back_to_generate_other_file():
    global mappings,bool1
    global selected_file ,var_file,var_type_of_file
    global mappings_supplier, mappings_supplier_address, mappings_supplier_site, mappings_supplier_third_party_relationship, mappings_supplier_site_assignment, mappings_supplier_contact, mappings_supplier_contact_address, mappings_supplier_profile_attachment, mappings_supplier_site_attachment, mappings_business_class_attachment, mappings_business_classification, mappings_product_and_service_category, mappings_supplier_payee, mappings_supplier_bank_accounts, mappings_bank_account_assignment
    # options_supplier = ["Transform Supplier Profile", "Transform Supplier Address Profile", "Transform Supplier Site Profile" , "Transform Supplier Third \nParty Relationship Profile" , "Transform Supplier \nSite Assignment Profile" , "Transform Supplier Contact Profile" , "Transform Contact Address Profile" , "Transform Supplier Profile\n Attachment Profile" , "Transform Supplier Profile\n Attachment Profile" , "Transform Supplier Site\n Attachment Profile" , "Transform Supplier Business\n Class Attachment Profile" , "Transform Supplier\n Business Classification Profile" , "Transform Supplier Product and Service\n Category Profile" , "Transform Supplier Payee Profile" , "Transform Supplier Bank\n Accounts Profile","Transform Supplier Bank \nAccount Assignment Profile"  ]

    # print (selected_file)
    if bool1:  
        # print(mappings)
        # print(selected_file)
        if selected_file == "Transform Supplier Profile":
            mappings_supplier = mappings.copy()
        # do something with mappings_supplier
        elif selected_file == "Transform Supplier Address Profile":
            mappings_supplier_address = mappings.copy()
            # do something with mappings_supplier_address
        elif selected_file == "Transform Supplier Site Profile":
            mappings_supplier_site = mappings.copy()
            # do something with mappings_supplier_site
        elif selected_file == "Transform Supplier Third \nParty Relationship Profile":
            mappings_supplier_third_party_relationship = mappings.copy()
            # do something with mappings_supplier_address
        elif selected_file == "Transform Supplier \nSite Assignment Profile":
            mappings_supplier_site_assignment = mappings.copy()
            # do something with mappings_supplier_site
        elif selected_file == "Transform Supplier Contact Profile":
            mappings_supplier_contact = mappings.copy()
            # do something with mappings_supplier_address
        elif selected_file == "Transform Supplier Contact Address Profile":
            mappings_supplier_contact_address = mappings.copy()
            # do something with mappings_contact_address
        elif selected_file == "Transform Supplier Profile\n Attachment Profile":
            mappings_supplier_profile_attachment = mappings.copy()
            # do something with mappings_supplier_profile_attachment
        elif selected_file == "Transform Supplier Site\n Attachment Profile":
            mappings_supplier_site_attachment = mappings.copy()
            # do something with mappings_supplier_site_attachment
        elif selected_file == "Transform Supplier Business\n Class Attachment Profile":
            mappings_business_class_attachment = mappings.copy()
            # do something with mappings_business_class_attachment
        elif selected_file == "Transform Supplier\n Business Classification Profile":
            mappings_business_classification = mappings.copy()
            # do something with mappings_business_classification
        elif selected_file == "Transform Supplier Product and Service\n Category Profile":
            mappings_product_and_service_category = mappings.copy()
            # do something with mappings_product_and_service_category
        elif selected_file == "Transform Supplier Payee Profile":
            mappings_supplier_payee = mappings.copy()
            # do something with mappings_supplier_payee
        elif selected_file == "Transform Supplier Bank\n Accounts Profile":
            mappings_supplier_bank_accounts = mappings.copy()
            # do something with mappings_supplier_bank_accounts
        elif selected_file == "Transform Supplier Bank \nAccount Assignment Profile":
            mappings_bank_account_assignment = mappings.copy()
            # do something with mappings_bank_account_assignment


        
    bool1 = False
    mappings.clear() 
    # print(mappings_supplier)
    var_file.set("Select File Type")

    # print(mappings_supplier)

    description_mapping.place_forget()
    heading_label.place_forget()
    heading_label1.place_forget()
    frame_mapping.place_forget()
    # frame.place(x=610, y=245)
    frame.bind("<Configure>",  frame.place(relx=0.5, rely=0.5, anchor='center'))
    option_menu_supplier.place(x=25,y=150)
    btn_next1.place(x = 25, y =202 )
    change_state(btn_next1)
    btn_back1.place(x = 100 , y = 270)
    def enable_button1(*args):
    # Check if an option is selected
        if var_file.get():
            btn_next1.config(state=tk.NORMAL)
        else:
            btn_next1.config(state=tk.DISABLED)
        # print(mappings_supplier)

        # print( selected_file)
        # print(mappings)
    var_file.trace("w", enable_button1)

def home():
    global bool1,bool2,selected_file, selected_file_type,mappings,var_file,var_type_of_file,selected_file_mapping
    global mappings_project,mappings_supplier, mappings_supplier_address, mappings_supplier_site, mappings_supplier_third_party_relationship, mappings_supplier_site_assignment, mappings_supplier_contact, mappings_supplier_contact_address, mappings_supplier_profile_attachment, mappings_supplier_site_attachment, mappings_business_class_attachment, mappings_business_classification, mappings_product_and_service_category, mappings_supplier_payee, mappings_supplier_bank_accounts, mappings_bank_account_assignment

    bool1 = False
    bool2 = False
    selected_file = None
    selected_file_type = None
    selected_file_mapping = None
    mappings_project = {}
    mappings_supplier ={}
    mappings_supplier_address ={}
    mappings_supplier_site ={}
    mappings_supplier_third_party_relationship ={}
    mappings_supplier_site_assignment ={}
    mappings_supplier_contact ={} 
    mappings_supplier_contact_address ={}
    mappings_supplier_profile_attachment ={}
    mappings_supplier_site_attachment ={}
    mappings_business_class_attachment ={}
    mappings_business_classification ={}
    mappings_product_and_service_category ={}
    mappings_supplier_payee ={}
    mappings_supplier_bank_accounts ={}
    mappings_bank_account_assignment ={}
    mappings = {}
    description_mapping.place_forget()
    heading_label.place_forget()
    heading_label1.place_forget()
    frame_mapping.place_forget()
    btn_next1.place_forget()
    btn_back1.place_forget()
    option_menu_supplier.place_forget()
    supplier_consolidated.clear()
    data_list.clear()
    # frame.place(x=610, y=245)
    frame.bind("<Configure>",  frame.place(relx=0.5, rely=0.5, anchor='center'))
    btn_upload.place(x= 25, y = 202)
    option_menu_file.place(x= 25, y = 150)
    var_type_of_file.set("Upload Process")
    change_state(option_menu_file)
    change_state(btn_upload)
    def enable_button1(*args):
    # Check if an option is selected
        if var_type_of_file.get():
            btn_upload.config(state=tk.NORMAL)
        else:
            btn_upload.config(state=tk.DISABLED)

        # print( selected_file)
        # print(mappings)
    var_file.trace("w", enable_button1)
    # print(data_list)

def confirm():
    global bool2
    bool2 = True
    show_mapping_window()
    
def show_mapping_window():
    global bool1,bool2, mappings, selected_file,var_file_mapping,selected_file_mapping, selected_file_type
    global mappings_supplier, mappings_supplier_address, mappings_supplier_site, mappings_supplier_third_party_relationship, mappings_supplier_site_assignment, mappings_supplier_contact, mappings_supplier_contact_address, mappings_supplier_profile_attachment, mappings_supplier_site_attachment, mappings_business_class_attachment, mappings_business_classification, mappings_product_and_service_category, mappings_supplier_payee, mappings_supplier_bank_accounts, mappings_bank_account_assignment, project_task_detail,project_details

    # print(224, mappings_supplier)
    bool1 = True
    if bool2:
        selected_file = selected_file_mapping

    def generate_mapped_excel_file(mapped_headers, data_list, list_type):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            workbook = openpyxl.Workbook()
            sheet = workbook.active

            # Write the headers to the first row
            headers = list_type
            sheet.append(headers)
            target_column = 'Import Action'
            if target_column in list_type:
                column_index = list_type.index(target_column) +1
            # print(len(data_list[1]))
                for count in range (len(data_list[0][1])) :
                    sheet.cell(row=count + 2, column=column_index, value="Create")
            # Write the column data
            for column_heading, column_values in data_list:
                if column_heading in mapped_headers:
                    target_column = mapped_headers[column_heading]  # Get the mapped target column
                    column_index = list_type.index(target_column) + 1  # Get the index of the target column
                    for i, value in enumerate(column_values):
                        sheet.cell(row=i + 2, column=column_index, value=value)  # Write the value to the specific column

            # Save the workbook as an XLSX file
            workbook.save(file_path)
            messagebox.showinfo("EY DATA CONVERSION TOOL", "Save Successful!\n\nPlease press Next")



#Frame
    change_state(btn_next1)
    frame.place_forget()
    frame_mapping.place(x=0,y=0)
    # update frame position
    frame_mapping.bind("<Configure>", frame_mapping.place(relx=0.5, rely=0.5, anchor='center'))
    custom_font = font.Font(size=12)
    # btn cconfirm
    btn_confirm = bttn(frame_mapping, 100, -283, "Confirm", '#000000', '#E5E8E8','#525252',confirm,'disabled',None,1,12,("Calibri Light", 10))        
    btn_confirm.bind("<Configure>", btn_confirm.place(relx=0.5, rely=0.5, anchor='center'))
# File to Generate 
    # print(263,mappings_supplier)
    # print (265,mappings_supplier)   

    option_menu_supplier_mapping.place(x=475,y=55)
    other_list=[]
    generating_list=[]
    # print(selected_file)
    # print(269,mappings_supplier)
    if selected_file_type == "Supplier": 
        if selected_file == "Transform Supplier Profile":
            other_list = supplier_columns[:]
            mappings = mappings_supplier.copy()
            # do something with other_list and mappings
        elif selected_file == "Transform Supplier Address Profile":
            other_list = supplier_address[:]
            mappings = mappings_supplier_address.copy()
            # do something with other_list and mappings
        elif selected_file == "Transform Supplier Site Profile":
            mappings = mappings_supplier_site.copy()
            other_list = supplier_site_data[:]
            # do something with other_list and mappings
        elif selected_file == "Transform Supplier Third \nParty Relationship Profile":
            other_list = supplier_third_party_relationship[:]
            mappings = mappings_supplier_third_party_relationship.copy()
            # do something with other_list and mappings
        elif selected_file == "Transform Supplier \nSite Assignment Profile":
            mappings = supplier_site_assignment.copy()
            other_list = supplier_site_assignment[:]
            # do something with other_list and mappings
        elif selected_file == "Transform Supplier Contact Profile":
            mappings = mappings_supplier_contact.copy()
            other_list = supplier_contact[:]
            # do something with other_list and mappings
        elif selected_file == "Transform Supplier Contact Address Profile":
            mappings = mappings_supplier_contact_address.copy()
            other_list = supplier_contact_address[:]
            # do something with other_list and mappings
        elif selected_file == "Transform Supplier Profile\n Attachment Profile":
            mappings = mappings_supplier_profile_attachment.copy()
            other_list = supplier_profile_attachment[:]
            # do something with other_list and mappings
        elif selected_file == "Transform Supplier Site\n Attachment Profile":
            mappings = mappings_supplier_site_attachment.copy()
            other_list = supplier_site_attachment[:]
            # do something with other_list and mappings
        elif selected_file == "Transform Supplier Business\n Class Attachment Profile":
            mappings = mappings_business_class_attachment.copy()
            other_list = supplier_business_class_attachment[:]
            # do something with other_list and mappings
        elif selected_file == "Transform Supplier\n Business Classification Profile":
            mappings = mappings_business_classification.copy()
            other_list = supplier_business_classification[:]
            # do something with other_list and mappings
        elif selected_file == "Transform Supplier Product and Service\n Category Profile":
            mappings = mappings_product_and_service_category.copy()
            other_list = supplier_product_and_service_category[:]
            # do something with other_list and mappings
        elif selected_file == "Transform Supplier Payee Profile":
            mappings = mappings_supplier_payee.copy()
            other_list = supplier_payee[:]
            # do something with other_list and mappings
        elif selected_file == "Transform Supplier Bank\n Accounts Profile":
            mappings = mappings_supplier_bank_accounts.copy()
            other_list = supplier_bank_accounts[:]
            # do something with other_list and mappings
        elif selected_file == "Transform Supplier Bank \nAccount Assignment Profile":
            mappings = mappings_bank_account_assignment.copy()
            other_list = supplier_bank_account_assignment[:]
            # do something with other_list and mappings

                    # options_project = ["Project details","Task Deatils"]

    elif selected_file_type == "Project": 
        if selected_file == "Project details":
            other_list = project_details[:]
            mappings = mappings_project.copy()
            # do something with other_list and mappings
        elif selected_file == "Task Deatils":
            other_list = project_task_detail[:]
            mappings = mappings_task_details.copy()
            # do something with other_list and mappings
    var_file_mapping.set(selected_file)
    generating_list =other_list[:]
    bool2 = False
    # print (selected_file)
    def enable_button(*args):
        global selected_file_type, selected_file,var_file_mapping,mappings,mappings_supplier,mappings_supplier_address,mappings_supplier_site,bool2,selected_file_mapping
        global mappings_task_details, mappings_project, mappings_supplier, mappings_supplier_address, mappings_supplier_site, mappings_supplier_third_party_relationship, mappings_supplier_site_assignment, mappings_supplier_contact, mappings_supplier_contact_address, mappings_supplier_profile_attachment, mappings_supplier_site_attachment, mappings_business_class_attachment, mappings_business_classification, mappings_product_and_service_category, mappings_supplier_payee, mappings_supplier_bank_accounts, mappings_bank_account_assignment

        if ( selected_file_type == "Supplier"):
        # Check if an option is selected
            if (var_file_mapping.get()):
                btn_confirm.config(state=tk.NORMAL)
                if selected_file == "Transform Supplier Profile":
                    mappings_supplier = mappings.copy()
                    # print("mappings_supplier: ", mappings_supplier)
                elif selected_file == "Transform Supplier Address Profile":
                    mappings_supplier_address = mappings.copy()
                    # print("mappings_supplier_address: ", mappings_supplier_address)
                elif selected_file == "Transform Supplier Site Profile":
                    mappings_supplier_site = mappings.copy()
                    # print("mappings_supplier_site: ", mappings_supplier_site)
                elif selected_file == "Transform Supplier Third \nParty Relationship Profile":
                    mappings_supplier_third_party_relationship = mappings.copy()
                    # print("mappings_supplier_third_party_relationship: ", mappings_supplier_third_party_relationship)
                elif selected_file == "Transform Supplier \nSite Assignment Profile":
                    mappings_supplier_site_assignment = mappings.copy()
                    # print("mappings_supplier_site_assignment: ", mappings_supplier_site_assignment)
                elif selected_file == "Transform Supplier Contact Profile":
                    mappings_supplier_contact = mappings.copy()
                    # print("mappings_supplier_contact: ", mappings_supplier_contact)
                elif selected_file == "Transform Supplier Contact Address Profile":
                    mappings_supplier_contact_address = mappings.copy()
                    # print("mappings_supplier_contact_address: ", mappings_supplier_contact_address)
                elif selected_file == "Transform Supplier Profile\n Attachment Profile":
                    mappings_supplier_profile_attachment = mappings.copy()
                    # print("mappings_supplier_profile_attachment: ", mappings_supplier_profile_attachment)
                elif selected_file == "Transform Supplier Site\n Attachment Profile":
                    mappings_supplier_site_attachment = mappings.copy()
                    # print("mappings_supplier_site_attachment: ", mappings_supplier_site_attachment)
                elif selected_file == "Transform Supplier Business\n Class Attachment Profile":
                    mappings_business_class_attachment = mappings.copy()
                    # print("mappings_business_class_attachment: ", mappings_business_class_attachment)
                elif selected_file == "Transform Supplier\n Business Classification Profile":
                    mappings_business_classification = mappings.copy()
                    # print("mappings_business_classification: ", mappings_business_classification)
                elif selected_file == "Transform Supplier Product and Service\n Category Profile":
                    mappings_product_and_service_category = mappings.copy()
                    # print("mappings_product_and_service_category: ", mappings_product_and_service_category)
                elif selected_file == "Transform Supplier Payee Profile":
                    mappings_supplier_payee = mappings.copy()
                    # print("mappings_supplier_payee: ", mappings_supplier_payee)
                elif selected_file == "Transform Supplier Bank\n Accounts Profile":
                    mappings_supplier_bank_accounts = mappings.copy()
                    # print("mappings_supplier_bank_accounts: ", mappings_supplier_bank_accounts)
                elif selected_file == "Transform Supplier Bank \nAccount Assignment Profile":
                    mappings_bank_account_assignment = mappings.copy()
                    # print("mappings_bank_account_assignment: ", mappings_bank_account_assignment)

            else:
                btn_confirm.config(state=tk.DISABLED)
            selected_file_mapping = var_file_mapping.get()
            # print(selected_file_mapping)
            # options_project = ["Project details","Task Deatils"]
        elif ( selected_file_type == "Project"):
            if (var_file_mapping.get()):
                btn_confirm.config(state=tk.NORMAL)
                if selected_file == "Task Deatils":
                    mappings_task_details = mappings.copy()
                    # print("mappings_supplier: ", mappings_supplier)
                elif selected_file == "Project details":
                    mappings_project = mappings.copy()
            
        var_file_mapping.trace("w", enable_button)

        
#Drop function
    def start_drag(event):
        index = listbox_supplier_consolidated.nearest(event.y)
        dragged_value = listbox_supplier_consolidated.get(index)
        listbox_supplier_consolidated.selection_set(index)
        listbox_supplier_consolidated._drag_data = dragged_value
        listbox_supplier_consolidated._drag_index = index
        listbox_supplier_consolidated.bind("<B1-Motion>", on_drag)

    def on_drag(event):
        listbox_supplier_consolidated.selection_clear(0, tk.END)
        index = listbox_supplier_consolidated.nearest(event.y)
        listbox_supplier_consolidated.selection_set(listbox_supplier_consolidated._drag_index, index)

    def on_drop(event):
        
        data = listbox_supplier_consolidated._drag_data
        index = listbox_other_file.nearest(event.y)
        destination_value = listbox_other_file.get(index)
        # print (data)
        # if (destination_value == "CLEAR MAPPING FOR THIS" and (data in mappings) ):
        #     del mappings[data]
            # print(mappings)
        # elif ( destination_value != "CLEAR MAPPING FOR THIS" ):
        mappings[data] = destination_value
            # print(mappings)
        # else: 
        #     return
        # Clean up
        listbox_supplier_consolidated.unbind("<B1-Motion>")
        listbox_supplier_consolidated.selection_clear(0, tk.END)
        listbox_supplier_consolidated._drag_data = None
        listbox_supplier_consolidated._drag_index = None
        update_mappings_display()
        
    def clear_mapping_for_selected():
        data= listbox_supplier_consolidated._drag_data
        if (data in mappings):
            del mappings[data]
        update_mappings_display()

#Listboxes
    # Central MAPPING description-
    def update_mappings_display():
        central_listbox.delete(0, tk.END)  # Clear previous mappings
        # print ( "I am here" )
        for element1, element2 in mappings.items():
            mapping_text1 = f"{element1} => {element2}"
            central_listbox.insert(tk.END,mapping_text1) 
    description_mapping.place(x=0 , y=-323)

    central_listbox = tk.Listbox(frame_mapping, height=20,width= 40,exportselection=False,bd=5)
    central_listbox.place(x=0, y=-100)

    central_listbox_scrollbar = tk.Scrollbar(central_listbox)
    central_listbox_scrollbar.place(relx=1,rely=0,relheight=1,anchor=tk.NE)
    central_listbox_scrollbar.configure(command=central_listbox.yview)

    def update_description_mapping(event):
        description_mapping.place(relx=0.5, rely=0.5, anchor='center')
    description_mapping.bind("<Configure>", update_description_mapping)

    def update_central_listbox_position(event):
            central_listbox.place(relx=0.5, rely=0.5, anchor='center')
    central_listbox.bind("<Configure>", update_central_listbox_position)

    # listbox_supplier_consolidated
    heading_label.place(x=-409 , y=-277)

    heading_label1.place(x=400, y=-277)
    
    listbox_supplier_consolidated = tk.Listbox(frame_mapping, height=31,width= 46,exportselection=False,font = custom_font,bd=5,borderwidth=5)
    listbox_supplier_consolidated.place(x=-408, y=34)

    scrollbar_supplier_consolidated = tk.Scrollbar(listbox_supplier_consolidated)
    scrollbar_supplier_consolidated.place(relx=1,rely=0,relheight=1,anchor=tk.NE)

    supplier_consolidated_list = supplier_consolidated[:]
    for item in supplier_consolidated_list:
                listbox_supplier_consolidated.insert(tk.END, item )
    
    listbox_supplier_consolidated.configure(yscrollcommand=scrollbar_supplier_consolidated.set)
    scrollbar_supplier_consolidated.configure(command=listbox_supplier_consolidated.yview)
    frame_mapping.bind("<Configure>",frame_mapping.place(relx=0.5, rely=0.5, anchor='center'))
    
    # Other ListBox
    listbox_other_file = tk.Listbox(frame_mapping, height=29,width= 46,exportselection=False,font = custom_font,bd=5)
    listbox_other_file.place(x=400, y=50)

    scrollbar_other_file = tk.Scrollbar(listbox_other_file)
    scrollbar_other_file.place(relx=1,rely=0,relheight=1,anchor=tk.NE)
    # "Generate Supplier File", "Generate Address File", "Generate Supplier Site File"
    
    for item in other_list:
                listbox_other_file.insert(tk.END, item )
    
    listbox_other_file.configure(yscrollcommand=scrollbar_other_file.set)
    scrollbar_other_file.configure(command=listbox_other_file.yview)
    # frame_mapping.bind("<Configure>",frame_mapping.place(relx=0.5, rely=0.5, anchor='center'))
    #Allignment of list boxes 
    def update_listbox1_position(event):
        listbox_supplier_consolidated.place(relx=0.5, rely=0.5, anchor='center')

    def update_listbox1_heading_position(event):
        heading_label.place(relx=0.5, rely=0.5, anchor='center')

    def update_listbox2_position(event):
        listbox_other_file.place(relx=0.5, rely=0.5, anchor='center')

    def update_listbox2_heading_position(event):
        heading_label1.place(relx=0.5, rely=0.5, anchor='center')
           
    listbox_supplier_consolidated.bind("<Configure>", update_listbox1_position)
    heading_label.bind("<Configure>", update_listbox1_heading_position)
    listbox_other_file.bind("<Configure>", update_listbox2_position)
    heading_label1.bind("<Configure>", update_listbox2_heading_position)

#Drag and Map    
    listbox_supplier_consolidated.bind("<Button-1>", start_drag)
    listbox_other_file.bind("<Button-1>", on_drop)
    update_mappings_display()

    def clear_mapping(): 
        mappings.clear()
        update_mappings_display()

    def save_mapping():
        save_dict_as_md(mappings)
    
    def load_mappings():
        global mappings
        value = load_dict_from_md()
        if value == None:
            return
        else:
            mappings = value.copy()
        update_mappings_display()
    


    
#Buttons

    # save dict button 

    btn_save_dict = bttn(frame_mapping, -100, 100, "Save Mapping", '#000000', '#E5E8E8','#525252',save_mapping,'normal',None,2,13,("Calibri Light", 10))        
    btn_save_dict.bind("<Configure>", btn_save_dict.place(relx=0.5, rely=0.5, anchor='center'))
    
    #loaddict button 
    btn_save_dict = bttn(frame_mapping, 100, 100, " Load Mapping", '#000000', '#E5E8E8','#525252',load_mappings,'normal',None,2,13,("Calibri Light", 10))        
    btn_save_dict.bind("<Configure>", btn_save_dict.place(relx=0.5, rely=0.5, anchor='center'))

    # clear mapping button
    btn_clear_mapping = bttn(frame_mapping, 0, 160, "Clear all the mappings", '#000000', '#E5E8E8','#525252',clear_mapping,'normal',None,2,27,("Calibri Light", 10))        
    btn_clear_mapping.bind("<Configure>", btn_clear_mapping.place(relx=0.5, rely=0.5, anchor='center'))
    # heading_label1 = tk.Label(window, text="Target",bg='#012E5F',fg="#FFFFFF", font=("Arial", 14,"bold"),height=1,width=35)
    # clear mapping for seletected button
    btn_clear_mappin_for_selected = bttn(frame_mapping, 400, -250, "Clear Mapping For This", '#FFFFFF', '#000000','#EEF2F7',clear_mapping_for_selected,'normal',None,1,42,("Calibri Light", 13,"bold"))
    btn_clear_mappin_for_selected.bind("<Configure>", btn_clear_mappin_for_selected.place(relx=0.5, rely=0.5, anchor='center'))
    # Export excell file button 
    btn_generate_xlsx = bttn(frame_mapping, 0, 210, "Generate Excel File", '#000000', '#E5E8E8','#525252',lambda: generate_mapped_excel_file(mappings, data_list, generating_list),'normal',photo,30,190,("Calibri Light", 10))
    btn_generate_xlsx.bind("<Configure>", btn_generate_xlsx.place(relx=0.5, rely=0.5, anchor='center'))
    #back button page 3 
    btn_generate_other_file  = bttn(frame_mapping, 0, 260, "Generate Other File", '#000000', '#E5E8E8','#525252',back_to_generate_other_file,'normal',None,2,27,("Calibri Light", 10))
    btn_generate_other_file.bind("<Configure>", btn_generate_other_file.place(relx=0.5, rely=0.5, anchor='center'))
    #btn home
    btn_home  = bttn(frame_mapping, 0, 310, "Home", '#000000', '#E5E8E8','#525252',home,'normal',None,2,27,("Calibri Light", 10))        
    btn_home.bind("<Configure>", btn_home.place(relx=0.5, rely=0.5, anchor='center'))
    # print(mappings)

# Create the main window
window = tk.Tk()
bool1 = False
bool2 = False
mappings_task_details = {}
mappings_project = {}
mappings_supplier ={}
mappings_supplier_address ={}
mappings_supplier_site ={}
mappings_supplier_third_party_relationship ={}
mappings_supplier_site_assignment ={}
mappings_supplier_contact ={}
mappings_supplier_contact_address ={}
mappings_supplier_profile_attachment ={}
mappings_supplier_site_attachment ={}
mappings_business_class_attachment ={}
mappings_business_classification ={}
mappings_product_and_service_category ={}
mappings_supplier_payee ={}
mappings_supplier_bank_accounts ={}
mappings_bank_account_assignment ={}
mappings = {}
selected_file = None
selected_file_type = None
selected_file_mapping = None
def enable_button(*args):
    global selected_file_type
    # Check if an option is selected
    if var_type_of_file.get():
        btn_upload.config(state=tk.NORMAL)
    else:
        btn_upload.config(state=tk.DISABLED)
    selected_file_type = var_type_of_file.get()

window.title("EY DATA CONVERSION TOOL")
window.geometry("400x300")
window.minsize(400, 300)

# UI elements

#images

#background
background_image = Image.open("backgroud.jpg")
custom_width = window.winfo_screenwidth()
custom_height = window.winfo_screenheight()
resized_image = background_image.resize((custom_width, custom_height))
background_image = ImageTk.PhotoImage(resized_image)
background_label = tk.Label(window, image=background_image)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

# frame
frame = customtkinter.CTkFrame(window, width=310, height=310,corner_radius=15, border_width=0 ,fg_color='#DEDEDE',bg_color='#3a3a3a')
frame.place(x=0, y=0)

frame_mapping = customtkinter.CTkFrame(window, width=1300, height=700,corner_radius=15, border_width=0 ,fg_color='#DEDEDE',bg_color='#3a3a3a')
frame_mapping.place_forget()

#EY_logo
image_ey = Image.open("ey-logo-black.jpg")
image_ey = image_ey.resize((60, 45))  # Resize the image if needed
image_eyy = ImageTk.PhotoImage(image_ey)
label_ey = tk.Label(frame, image=image_eyy)
label_ey.place(x=120,y=35) 

#Excel logo
image = Image.open("Microsoft_Office_Excel.png")
image = image.resize((20, 20))  # Resize the image if needed
photo = ImageTk.PhotoImage(image)

#labels 

lbl_file_gen_tool = tk.Label(frame, text="Oracle Cloud FBDI \nFile Generator", bg='#DEDEDE', fg="Black",  font=("Arial", 15,"bold"))
lbl_file_gen_tool.place(x= 60,y=90)
description_mapping = tk.Label(window, text="""Kindly map the fields from source file to Target FBDI file and Click on generate button  """,bg='#012E5F',fg="#FFFFFF", font=("Arial", 10),width=162)
description_mapping.place_forget()
heading_label = tk.Label(window, text="Source",bg='#012E5F',fg="#FFFFFF", font=("Arial", 14,"bold"),height=1,width=35)
heading_label.place_forget()
heading_label1 = tk.Label(window, text="Target",bg='#012E5F',fg="#FFFFFF", font=("Arial", 14,"bold"),height=1,width=35)
heading_label.place_forget()
#buttons 
btn_upload= bttn(frame, 25, 220, "Upload Excel File  ", '#FDFEFE', '#000000','#E5E8E8', upload_excel,'disabled',photo,30,255,("Calibri Light", 10))
btn_next =bttn(frame, 25, 270, "Next", '#808080', '#FDFEFE','#17202A', show_type_of_files,'normal',"" ,1,10,("Calibri Light", 11,'bold'))
btn_next1 =bttn(frame, 30, 202, "Next",'black', '#FFFFFF',"#525252", show_mapping_window,'disabled',None,1,28,("Calibri Light", 11,'bold'))
btn_back1 =bttn(frame, 25, 270, "Back", '#2C3E50', '#FDFEFE','#17202A', back1,'normal',None ,1,10,("Calibri Light", 11,'bold'))
btn_next1.place_forget()
btn_back1.place_forget()
btn_next.place_forget()

#Option menu's 

#file_type
var_type_of_file = tk.StringVar(window)
var_type_of_file.set("Upload Process")
options = ["Supplier", "Customer","Project"]
option_menu_file = opt_menu(frame, 25, 170, 1, 1, var_type_of_file ,options,  '#FDFEFE', '#000000','#E5E8E8', None,'normal',40,37 )
var_type_of_file.trace("w", enable_button)

#Supplier file 
var_file = tk.StringVar(window)
var_file.set("Select File Type")
options_supplier = ["Transform Supplier Profile", "Transform Supplier Address Profile", "Transform Supplier Site Profile" , "Transform Supplier Third \nParty Relationship Profile" , "Transform Supplier \nSite Assignment Profile" , "Transform Supplier Contact Profile" , "Transform Supplier Contact Address Profile" , "Transform Supplier Profile\n Attachment Profile" , "Transform Supplier Site\n Attachment Profile" , "Transform Supplier Business\n Class Attachment Profile" , "Transform Supplier\n Business Classification Profile" , "Transform Supplier Product and Service\n Category Profile" , "Transform Supplier Payee Profile" , "Transform Supplier Bank\n Accounts Profile","Transform Supplier Bank \nAccount Assignment Profile"  ]
option_menu_supplier = opt_menu(frame, 25, 150, 1, 1, var_file,options_supplier,  '#FDFEFE', '#000000','#E5E8E8', None,'normal', 40 ,37 )
option_menu_supplier.place_forget()


var_file_proj = tk.StringVar(window)
var_file_proj.set("Select File Type")
options_project = ["Project details","Task Deatils"]
option_menu_project = opt_menu(frame, 25, 150, 1, 1, var_file,options_project,  '#FDFEFE', '#000000','#E5E8E8', None,'normal', 40 ,37 )
option_menu_project.place_forget()

var_file_mapping = tk.StringVar(window)
option_menu_supplier_mapping = opt_menu(frame_mapping, 25, 150, 1, 1, var_file_mapping,options_supplier,  '#FDFEFE', '#000000','#E5E8E8', None,'normal', 35 ,30 )
option_menu_supplier_mapping.place_forget()
frame.bind("<Configure>",  frame.place(relx=0.5, rely=0.5, anchor='center'))

# Start the main event loop
window.mainloop()
