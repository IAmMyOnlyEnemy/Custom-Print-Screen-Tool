import tkinter
import customtkinter as ctk
import json
import pyautogui
import os
import win32clipboard
from tkinter.filedialog import askdirectory
from io import BytesIO

numbers_vals = ["{:02d}".format(x) for x in range(1,11)]

letters_vals = [chr(x) for x in range(65,75)]

# Create a dictionary with settings imported from a json file or default if there is no json:
if os.path.exists("./"+"configuration.json"):
    with open('configuration.json','r') as json_file:
        settings_dictionary = json.load(json_file)
else:
    settings_dictionary = dict({
                            'configs': [
                                        [250,150,750,650,"Config-1"],
                                        [200,140,760,640,"Config-2"],
                                        [300,130,770,630,"Config-3"],
                                        [350,120,780,620,"Config-4"]
                                        ],
                            'img_list': ["IMG0","IMG1","IMG2","IMG3","IMG4"],
                            'active_config': 0,
                            'dark_theme': 1,
                            'on_top': 0,
                            'clipboard': 0,
                            'work_path': os.getcwd(),
                            })

img_vals = settings_dictionary["img_list"]

sWdth, sHght = pyautogui.size()

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configure window
        self.title("Enhanced PrintScreen Tool")
        self.geometry(f"{500}x{600}")

        # Configure grid layout (4x1)
        self.grid_rowconfigure((0, 3), weight=0)
        self.grid_rowconfigure((1, 2), weight=1)

        # Create top bar frame with widgets. It will contain:
        #   - 1 Entry for folder path 
        #   - 1 Button for browse
        self.frame_browse = ctk.CTkFrame(self, corner_radius=5, width=500)
        self.frame_browse.grid(row=0, column=0, sticky="w")
        self.path_var = tkinter.StringVar(value=settings_dictionary["work_path"])
        self.entry_path = ctk.CTkEntry(self.frame_browse,textvariable=self.path_var,state="disabled",width=385)
        self.entry_path.grid(row=0, column=0, padx=(5, 0),pady=(5, 5), sticky="w")
        self.button_browse = ctk.CTkButton(self.frame_browse,text="Browse",width=100,command=self.browse_button_event)
        self.button_browse.grid(row=0, column=1, padx=(5, 5), pady=(5, 5), sticky="e")

    # ------------------------------------------------------------------------------------------ #
    #                                                                                            #
    #                                          TABS FRAME:                                       #
    #                                                                                            #
    # ------------------------------------------------------------------------------------------ #
        self.frame_tabview = ctk.CTkFrame(self, corner_radius=5)
        self.frame_tabview.grid(row=1, column=0, sticky="w")
        self.tabview_imgs = ctk.CTkTabview(self.frame_tabview, width=490, height=380)
        self.tabview_imgs.grid(row=0, column=0, padx=(5, 5), pady=(0, 5), sticky="nsew")
        tab_names = ["Single image",
                     "Multiple images",
                     "Settings"]

    #           #------------------------------------------------------------------------------- #
    #           #                                                                                #
    #           #                  ------  1ST TAB DEFINITION  ------                            #
    #           #                                                                                #
    #           #------------------------------------------------------------------------------- #
        self.tabview_imgs.add(tab_names[0])
        self.tabview_imgs.tab(tab_names[0]).grid_rowconfigure((0,1,2), weight=0)
        self.tabview_imgs.tab(tab_names[0]).grid_rowconfigure((3,4,5,6,7), weight=1)

        # 1st row:
        self.label_img_name = ctk.CTkLabel(master=self.tabview_imgs.tab(tab_names[0]), 
                            text="", font=ctk.CTkFont(size=12, weight="bold"))#,width=200)
        self.label_img_name.grid(row=0, column=0, columnspan=5, padx=(10, 0), pady=(5, 5), sticky="snew")

        # 2nd row:
        self.tabview_imgs_switches_var = []
        self.tabview_imgs_switches = []

        for i in range(4):
            switch_var = tkinter.IntVar(value=0)
            self.tabview_imgs_switches_var.append(switch_var)
            switch = ctk.CTkSwitch(master=self.tabview_imgs.tab(tab_names[0]),text="",width=40, 
                                    variable=self.tabview_imgs_switches_var[i],onvalue=1,offvalue=0)
            if i > 1:
                switch.grid(row=2,column=i+1,padx=(15,5),pady=(5,5),sticky="nwes")    
            else:
                switch.grid(row=2,column=i,padx=(15,5),pady=(5,5),sticky="nwes")
            self.tabview_imgs_switches.append(switch)

        # 3rd row:
        self.test_var = tkinter.StringVar(value=numbers_vals[0])
        self.optionemenu_test = ctk.CTkOptionMenu(master=self.tabview_imgs.tab(tab_names[0]),command=self.OM_event, 
                                    values=numbers_vals[:10],variable=self.test_var,width=50)
        self.optionemenu_test.grid(row=1,column=0,padx=(5,5),pady=(5,5),sticky="wn")

        self.subtest_var = tkinter.StringVar(value=letters_vals[0])
        self.optionemenu_subtest = ctk.CTkOptionMenu(master=self.tabview_imgs.tab(tab_names[0]),command=self.OM_event, 
                                    values=letters_vals,variable=self.subtest_var,width=50)
        self.optionemenu_subtest.grid(row=1,column=1,padx=(5,5),pady=(5,5),sticky="wn")

        self.iter_var = tkinter.StringVar(value=numbers_vals[0])
        self.optionemenu_iter = ctk.CTkOptionMenu(master=self.tabview_imgs.tab(tab_names[0]),command=self.OM_event, 
                                    values=numbers_vals,variable=self.iter_var,width=50)
        self.optionemenu_iter.grid(row=1,column=4,padx=(5,5),pady=(5,5),sticky="wn")

        #4th row:
        self.label_line = ctk.CTkLabel(master=self.tabview_imgs.tab(tab_names[0]),text="")
        self.label_line.grid(row=3, column=0, padx=(5, 5), pady=(5, 5), sticky="nw")

        self.img_name_var = tkinter.StringVar(value="")
        self.new_img_name_entry = ctk.CTkEntry(master=self.tabview_imgs.tab(tab_names[0]),
                                                placeholder_text="-",textvariable=self.img_name_var,
                                                width=120,height=20)
        self.new_img_name_entry.grid(row=4,column=0,columnspan=2,padx=(5, 5),pady=(5, 5),sticky="snew")
        self.new_img_name_entry.bind("<Return>", self.new_img_event)
        self.img_name_var.set(img_vals[0])

        self.button_up = ctk.CTkButton(master=self.tabview_imgs.tab(tab_names[0]),text="Up",width=40,
                                            command=self.up_button_event)
        self.button_up.grid(row=5, column=0, padx=(5, 5), pady=(5, 5), sticky="snew")
        self.button_down = ctk.CTkButton(master=self.tabview_imgs.tab(tab_names[0]),text="Down",width=40,
                                            command=self.down_button_event)
        self.button_down.grid(row=5, column=1, padx=(5, 5), pady=(5, 5), sticky="snew")
        self.button_delete = ctk.CTkButton(master=self.tabview_imgs.tab(tab_names[0]),text="Delete",width=100,
                                            command=self.delete_button_event)
        self.button_delete.grid(row=6, column=0, columnspan=2, padx=(5, 5), pady=(5, 5), sticky="snew")

        for i in range(2):
            self.label_line = ctk.CTkLabel(master=self.tabview_imgs.tab(tab_names[0]),text="")
            self.label_line.grid(row=6+i, column=0, padx=(5, 5), pady=(5, 5), sticky="w")

        self.frame_img_list = ctk.CTkScrollableFrame(master=self.tabview_imgs.tab(tab_names[0]),width=180)
        self.frame_img_list.grid(row=1,rowspan=8,column=2,padx=(5, 5),pady=(5, 5),sticky="snew")

        self.img_var_list = tkinter.StringVar(value=img_vals[0])
        self.imgs_radios = []
        self.fill_img_radios()

    #           #------------------------------------------------------------------------------- #
    #           #                                                                                #
    #           #                  ------  2ND TAB DEFINITION  ------                            #
    #           #                                                                                #
    #           #------------------------------------------------------------------------------- #
        self.tabview_imgs.add(tab_names[1])
        self.tabview_imgs.tab(tab_names[1]).grid_columnconfigure((0,1), weight=1)
        self.tabview_imgs.tab(tab_names[1]).grid_rowconfigure((0,1), weight=0)

        plchld_str = "For each member from the left column, do the actions from the right column"

        self.entry_prints = ctk.CTkLabel(master=self.tabview_imgs.tab(tab_names[1]),text_color="grey",
                                         text=plchld_str,corner_radius=5,justify="center")
        self.entry_prints.grid(row=0,column=0,columnspan=2,padx=(5,5),pady=(5,5),sticky="nsew")

        self.tb_subjects = ctk.CTkTextbox(master=self.tabview_imgs.tab(tab_names[1]),corner_radius=5,height=285)
        self.tb_subjects.grid(row=1,column=0,padx=(5,5),pady=(5,5),sticky="snew")
        self.tb_subjects.insert("0.0", "Some subjects!\n" * 5)

        self.tb_actions = ctk.CTkTextbox(master=self.tabview_imgs.tab(tab_names[1]),corner_radius=5,height=285)
        self.tb_actions.grid(row=1,column=1,padx=(5,5),pady=(5,5),sticky="snew")
        self.tb_actions.insert("0.0", "Some verbs!\n" * 20)

    #           #------------------------------------------------------------------------------- #
    #           #                                                                                #
    #           #                  ------  3RD TAB DEFINITION  ------                            #
    #           #                                                                                #
    #           #------------------------------------------------------------------------------- #
        self.tabview_imgs.add(tab_names[2])
        info_labels = ["Label","x1","y1","width","height","Description"]

        for i in range(6):
            self.label_config = ctk.CTkLabel(master=self.tabview_imgs.tab(tab_names[2]), 
                                               text=info_labels[i], font=ctk.CTkFont(size=12, weight="bold"))
            self.label_config.grid(row=0, column=i, padx=(5, 5), pady=(5, 5), sticky="snew")

        self.conf_desc_vals = []
        self.config_entries = [[],[],[],[]]

        for i in range(4):
            i_row = 2*i+1
            self.label_config = ctk.CTkLabel(master=self.tabview_imgs.tab(tab_names[2]),width=80, 
                                               text="Configure " + str(i+1) + ":", height=15,
                                               font=ctk.CTkFont(size=12, weight="bold"))
            self.label_config.grid(row=i_row, column=0, padx=(5, 5), pady=(5, 5), sticky="w")

            self.x1_val = tkinter.StringVar(value=settings_dictionary['configs'][i][0])
            self.x1_config = ctk.CTkEntry(master=self.tabview_imgs.tab(tab_names[2]),
                                                    textvariable=self.x1_val,width=45,height=15,)
            self.x1_config.grid(row=i_row, column=1, padx=(5, 5),pady=(5, 5), sticky="snew")

            self.y1_val = tkinter.StringVar(value=settings_dictionary['configs'][i][1])
            self.y1_config = ctk.CTkEntry(master=self.tabview_imgs.tab(tab_names[2]),
                                                    textvariable=self.y1_val,width=45,height=15,)
            self.y1_config.grid(row=i_row, column=2, padx=(5, 5),pady=(5, 5), sticky="snew")

            self.width_val = tkinter.StringVar(value=settings_dictionary['configs'][i][2])
            self.width_config = ctk.CTkEntry(master=self.tabview_imgs.tab(tab_names[2]),
                                                    textvariable=self.width_val,width=45,height=15,)
            self.width_config.grid(row=i_row, column=3, padx=(5, 5),pady=(5, 5), sticky="snew")

            self.height_val = tkinter.StringVar(value=settings_dictionary['configs'][i][3])
            self.height_config = ctk.CTkEntry(master=self.tabview_imgs.tab(tab_names[2]),
                                                    textvariable=self.height_val,width=45,height=15,)
            self.height_config.grid(row=i_row, column=4, padx=(5, 5),pady=(5, 5), sticky="w")

            self.description_val = tkinter.StringVar(value=settings_dictionary['configs'][i][4])
            self.description_config = ctk.CTkEntry(master=self.tabview_imgs.tab(tab_names[2]),
                                                    textvariable=self.description_val,width=150,height=15,)
            self.description_config.grid(row=i_row, column=5, padx=(5, 5),pady=(5, 5), sticky="w")
            self.conf_desc_vals.append(self.description_val)

            self.config_entries[i].append(self.x1_val)
            self.config_entries[i].append(self.y1_val)
            self.config_entries[i].append(self.width_val)
            self.config_entries[i].append(self.height_val)
            self.config_entries[i].append(self.description_val)

            res_val = "{0}x{1}[{2}x{3}] ".format(sWdth, 
                                                 sHght, 
                                                 settings_dictionary['configs'][i][2], 
                                                 settings_dictionary['configs'][i][3])
            self.label_params = ctk.CTkLabel(master=self.tabview_imgs.tab(tab_names[2]), 
                                             width=80, 
                                             text=res_val,height=15, 
                                             font=ctk.CTkFont(size=12,weight="bold",slant="italic"))
            self.label_params.grid(row=i_row+1, column=1, columnspan=4, padx=(5, 5), pady=(5, 5), sticky="n")

        self.label_info_text = "{0}\n{1}\n{2}\n{3}".format("x1      : number of pixels from left side of the screen",
                                                           "y1      : number of pixels from top side of the screen",
                                                           "width : width of print screen rectangle",
                                                           "height: height of print screen rectangle")
        self.label_info_params = ctk.CTkLabel(master=self.tabview_imgs.tab("Settings"), 
                                                text=self.label_info_text,
                                                font=ctk.CTkFont(size=12, slant="italic"), justify="left")
        self.label_info_params.grid(row=10, column=0, columnspan=6, padx=(5, 5), pady=(5, 5), sticky="w")

    # ------------------------------------------------------------------------------------------ #
    #                                                                                            #
    #                                      PARAMETERS FRAME:                                     #
    #                                                                                            #
    # ------------------------------------------------------------------------------------------ #
        self.frame_params = ctk.CTkFrame(self, corner_radius=0, width=500)
        self.frame_params.grid(row=2, column=0, sticky="w")
        self.frame_params.grid_columnconfigure((0, 1, 2), weight=1)
        self.frame_params.grid_rowconfigure((0, 1, 2, 3), weight=0)

        # Create radio buttons for configuration options
        self.radio_var = tkinter.IntVar(value=settings_dictionary['active_config'])
        self.radio_config_list = []
        for i in range(4):
            self.radio_config = ctk.CTkRadioButton(master=self.frame_params, variable=self.radio_var, 
                                                   value=i, text="Configure " + str(i+1), width=150)
            self.radio_config.grid(row=i, column=0, padx=(5, 5), pady=(5, 5), sticky="w")
            self.radio_config_list.append(self.radio_config)

        # Create switches
        self.theme_var = tkinter.IntVar(value=settings_dictionary['dark_theme'])
        self.switch_theme = ctk.CTkSwitch(master=self.frame_params, text="Dark theme",variable=self.theme_var, 
                                        onvalue=1, offvalue=0,command=self.theme_switch_event, width=160)
        self.switch_theme.grid(row=0, column=1, padx=(5,5), pady=(5, 5), sticky="w")

        self.on_top_var = tkinter.IntVar(value=settings_dictionary['on_top'])
        self.switch_ontop = ctk.CTkSwitch(master=self.frame_params, text="On top",variable=self.on_top_var,
                                        onvalue=1, offvalue=0,command=self.ontop_switch_event)
        self.switch_ontop.grid(row=1, column=1, padx=(5,5), pady=(5, 5), sticky="w")

        self.clipboard_var = tkinter.IntVar(value=settings_dictionary['clipboard'])
        self.switch_copyclip = ctk.CTkSwitch(master=self.frame_params, text="Copy to Clipboard",variable=self.clipboard_var,
                                        onvalue=1, offvalue=0,command=self.clipboard_switch_event)
        self.switch_copyclip.grid(row=2, column=1, padx=(5,5), pady=(5, 5), sticky="w")
        # Create the Settings button
        self.button_save = ctk.CTkButton(self.frame_params,text="Save configuration",command=self.save_button_event, 
                                          corner_radius=30,font=ctk.CTkFont(size=12, weight="bold"))
        self.button_save.grid(row=3, column=1, padx=(5, 5), pady=(5, 5), sticky="snew")
        # Create the Print button
        self.button_print = ctk.CTkButton(self.frame_params,text="Capture",command=self.print_button_event, 
                                          corner_radius=30, width=160,font=ctk.CTkFont(size=20, weight="bold"))
        self.button_print.grid(row=0, column=2,rowspan=4, padx=(5, 5), pady=(5, 5), sticky="ensw")

    # ------------------------------------------------------------------------------------------ #
    #                                                                                            #
    #                                        MESSAGE FRAME:                                      #
    #                                                                                            #
    # ------------------------------------------------------------------------------------------ #
        self.frame_message = ctk.CTkFrame(self, corner_radius=5)
        self.frame_message.grid(row=3, column=0, sticky="w")
        # Create labels:
        self.message_var = tkinter.StringVar(value="")
        self.label_message = ctk.CTkLabel(self.frame_message,textvariable=self.message_var,width=380,justify="left")
        self.label_message.grid(row=0, column=0, padx=(5, 5), pady=(5, 5), sticky="w")
        self.label_version = ctk.CTkLabel(self.frame_message, text="Version 0.9.9.9", width=100, justify="right")
        self.label_version.grid(row=0, column=1, padx=(5, 5), pady=(5, 5), sticky="e")

    # ------------------------------------------------------------------------------------------ #
    #                                                                                            #
    #                                     SET DEFAULT VALUES                                     #
    #                                                                                            #
    # ------------------------------------------------------------------------------------------ #
        self.tabview_imgs.set(tab_names[0])
        if self.theme_var.get() == 1:
            self.theme_switch_event()
        self.switch_ontop.deselect()
        self.switch_copyclip.deselect()
        self.update_label()
        self.update_config_desc()
        self.message_var.set("App ready")

    # ------------------------------------------------------------------------------------------ #
    #                                                                                            #
    #                             DEFINITION OF SPECIFIC FUNCTIONS                               #
    #                                                                                            #
    # ------------------------------------------------------------------------------------------ #
    def OM_event(self,event):
        # function for OptionMenu event 
        self.update_label()

    def new_img_event(self,event):
        # 1'st check if this value already exists:
        if self.img_name_var.get() in img_vals:
            self.message_var.set("Image name already exists")
        else:
            # insert the new name below current selefcted one:
            idx = img_vals.index(self.img_var_list.get())
            img_vals.insert(idx+1,self.img_name_var.get())
            self.fill_img_radios()
            self.img_var_list.set(self.img_name_var.get())
            self.message_var.set("New image name added")

    def up_button_event(self):
        # move name up in list
        idx = img_vals.index(self.img_name_var.get())
        if idx == 0:
            self.message_var.set("Already on top")
        else:
            img_vals.pop(idx)
            img_vals.insert(idx-1,self.img_name_var.get())
            self.fill_img_radios()
            self.message_var.set("Moved up")

    def down_button_event(self):
        # move name down in list
        idx = img_vals.index(self.img_name_var.get())
        if idx == len(img_vals)-1:
            self.message_var.set("Already on bottom")
        else:
            img_vals.pop(idx)
            img_vals.insert(idx+1,self.img_name_var.get())
            self.fill_img_radios()
            self.message_var.set("Moved down")

    def delete_button_event(self):
        # remove name from list
        if len(img_vals)>0:
            next_val = get_next(self.img_name_var.get(),img_vals)
            img_vals.remove(self.img_name_var.get())
            self.fill_img_radios()
            self.img_name_var.set(next_val)
            self.img_var_list.set(next_val)
            self.message_var.set("Image name removed")
        else:
            self.message_var.set("Nothing to remove")

    def fill_img_radios(self):
        # update the settings dictionary:
        settings_dictionary["img_list"] = img_vals

        # remove old radios buttons before updating them:
        for elems in self.imgs_radios:
            elems.destroy()
        self.imgs_radios.clear()

        # create new radios buttons according to the new list:
        i = 0
        for img_val in img_vals:
            self.radio_img = ctk.CTkRadioButton(master=self.frame_img_list, variable=self.img_var_list, 
                                                value=img_val,text=img_val,command=self.update_label)
            self.radio_img.grid(row=i, column=0, padx=(5, 5), pady=(5, 5), sticky="w")
            self.imgs_radios.append(self.radio_img)
            i += 1

    def browse_button_event(self):
        # function for browse button
        dir_path = askdirectory(title='Select Folder')

        if settings_dictionary["work_path"] == dir_path:
            self.message_var.set("Folder path was not changed")
        else:
            if dir_path != "":
                self.path_var.set(dir_path)
                settings_dictionary["work_path"] = dir_path
            self.message_var.set("Folder path was changed")
        self.update_label()

    def save_button_event(self):
        # 1st check that coordinates are numeric:
        self.message_var.set("Configuration was saved")
        for i in range(4):
            for j in range(5):
                if j < 4:
                    try:
                        settings_dictionary['configs'][i][j] = int(self.config_entries[i][j].get())
                    except:
                        self.message_var.set(f"Incorrect value at {i} : {j}")
                else:
                    settings_dictionary['configs'][i][j] = self.config_entries[i][j].get()
        self.update_config_desc()

        # save the configuration as json
        save_settings()

    def print_button_event(self):
        ps_x = settings_dictionary['configs'][self.radio_var.get()][0]
        ps_y = settings_dictionary['configs'][self.radio_var.get()][1]
        ps_w = settings_dictionary['configs'][self.radio_var.get()][2]
        ps_h = settings_dictionary['configs'][self.radio_var.get()][3]
        ps_name = self.label_img_name.cget("text")
        
        self.message_var.set(take_screenshot(ps_x,ps_y,ps_w,ps_h,ps_name))

        if self.tabview_imgs_switches_var[0].get() == 1:
            self.test_var.set(get_next(self.test_var.get(),self.optionemenu_test.cget("values")))
        if self.tabview_imgs_switches_var[1].get() == 1:
            self.subtest_var.set(get_next(self.subtest_var.get(),self.optionemenu_subtest.cget("values")))
        if self.tabview_imgs_switches_var[2].get() == 1:
            self.img_var_list.set(get_next(self.img_var_list.get(),img_vals))
        if self.tabview_imgs_switches_var[3].get() == 1:
            self.iter_var.set(get_next(self.iter_var.get(),self.optionemenu_iter.cget("values")))
        self.update_label()

    def theme_switch_event(self):
        settings_dictionary['dark_theme'] = self.switch_theme.get()
        if self.switch_theme.get() == 1:
            ctk.set_appearance_mode("Dark")
            ctk.set_default_color_theme("dark-blue")  # Themes: "blue" (standard), "green", "dark-blue"
            self.message_var.set("Dark theme activated")
        else:
            ctk.set_appearance_mode("Light")
            ctk.set_default_color_theme("green")  # Themes: "blue" (standard), "green", "dark-blue"
            self.message_var.set("Light theme activated")

    def ontop_switch_event(self):
        settings_dictionary['on_top'] = self.on_top_var.get()
        if self.on_top_var.get() == 1:
            # to do: always on top activated
            self.message_var.set("App will be on top")
            self.attributes('-topmost',True)
            self.update()
        else:
            # to do: always on top deactivated
            self.message_var.set("App will not be on top")
            self.attributes('-topmost',False)
            self.update()

    def clipboard_switch_event(self):
        settings_dictionary['clipboard'] = self.clipboard_var.get()
        if self.clipboard_var.get() == 1:
            # to do: copy to clipboard activated
            self.message_var.set("Image will be saved to clipboard")
        else:
            # to do: copy to clipboard deactivated
            self.message_var.set("Image will not be saved to clipboard")

    def update_label(self):
        self.label_img_name.configure(text=self.test_var.get() + self.subtest_var.get() + "_" + 
                                           self.img_var_list.get() + "_" + self.iter_var.get())
        self.img_name_var.set(self.img_var_list.get())

        if os.path.exists(settings_dictionary["work_path"]+"./"+self.label_img_name.cget("text")+".jpg"):
            self.label_img_name.configure(fg_color="red")
        else:
            self.label_img_name.configure(fg_color='transparent')

    def update_config_desc(self):
        for i in range(len(self.radio_config_list)):
            self.radio_config_list[i].configure(text=self.conf_desc_vals[i].get())



# ------------------------------------------------------------------------------------------ #
#                                                                                            #
#                             DEFINITION OF GENERAL FUNCTIONS                                #
#                                                                                            #
# ------------------------------------------------------------------------------------------ #
def get_next(my_val,my_arr):
    for i in range(len(my_arr)):
        if my_arr[i] == my_val:
            if i == len(my_arr)-1:
                return my_val
            return my_arr[i+1]
    if len(my_arr)>0:
        return my_arr[-1]
    else:
        return ""

def take_screenshot(ps_x,ps_y,ps_w,ps_h,ps_name):
    try:
        my_ps = pyautogui.screenshot(region=(ps_x,ps_y,ps_w,ps_h))
        my_ps.save(settings_dictionary['work_path']+'/'+ps_name+".jpg")
        clipboard_stat = ""
        if settings_dictionary['clipboard'] == 1:
            copy_img_to_clip(my_ps)
            clipboard_stat = " and added to clipboard."
        return "Image saved as {0}.jpg{1}".format(ps_name,clipboard_stat)
    except Exception as e:
        return "Error while saving: \ne: {0}\n{1};{2};{3};{4};{5}".format(e,ps_x,ps_y,ps_w,ps_h)

def save_settings():
    with open('configuration.json','w') as json_file:
        json_file.write(json.dumps(settings_dictionary))

def copy_img_to_clip(img = None):
    '''
        Copy the print screen to ClipBoard
    '''
    output = BytesIO()
    img.convert('RGB').save(output, 'BMP')
    data = output.getvalue()[14:]
    output.close()
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

# ------------------------------------------------------------------------------------------ #
#                                                                                            #
#                                    RUN THE MAIN PROGRAM                                    #
#                                                                                            #
# ------------------------------------------------------------------------------------------ #
if __name__ == "__main__":
    app = App()
    app.mainloop()