import tkinter as tk
import pandas as pd
import random

def unique_values(list):
    output = []
    for element in list:
        if element not in output:
            output.append(element)
    return output


class File:
    matrice_name = "matrice_source.xlsx"
    df_countries = pd.read_excel(matrice_name, sheet_name="Countries")
    df_therapyAreas = pd.read_excel(matrice_name, sheet_name="THERAPY_AREA_SETUP")


class ContentPlan(File):
    template_name = "fresh_template_content_plan.xlsx"

    def __init__(self, input):
        super().__init__()
        self.df_template = pd.read_excel(self.template_name)
        self.df_contentPlans = pd.read_excel(self.matrice_name, sheet_name="CONTENT_PLAN_SETUP")
        self.input = input

    def to_data_frame(self):
        content_plans   = self.df_contentPlans
        markets         = self.input.get("markets")
        securityProfile = self.input.get("securityProfile")
        userid          = self.input.get("MUDID")

        for iterator, market in enumerate(markets):
            if iterator == 0:
                tmp_df = content_plans[content_plans["Security_profile"] == securityProfile[0]]
                tmp_df = tmp_df.drop(columns=["Security_profile"])

            if iterator > 0:
                tmp_df = content_plans[content_plans["Security_profile"] == securityProfile[0]]
                tmp_df = tmp_df.drop(columns=["Security_profile"])
                tmp_df = tmp_df.dropna(subset=['country__c'])

            country_id = self.df_countries.loc[self.df_countries["Country Name"] == market]
            country_id_val = country_id.iloc[0, 4]

            tmp_df = tmp_df.replace({"{LOC}": country_id_val,
                                     "user_id": userid})

            self.df_template = self.df_template.append(tmp_df)

        return self.df_template


class InternalUserRoleSetup(File):
    template_name = "fresh_template_user_role_setup.xlsx"

    def __init__(self, input):
        super().__init__()

        self.df_template = pd.read_excel(self.template_name)
        self.df_roleSetup = pd.read_excel(self.matrice_name, sheet_name="ROLE_SETUP")
        self.df_specialRoles = pd.read_excel(self.matrice_name, sheet_name="SPECIAL_COUNTRY_ROLES")
        self.input = input
        self.therapyAreasRoles = ['Global Commercial Content Owner','Global Commercial Reviewer','Global Commercial Approver','Global Medical Content Owner','Global Medical Reviewer','Global Medical Approver']
        self.specialCountryRoles = ["Austria","Bulgaria","Canada","Costa Rica","Czech Republic","Egypt","El Salvador","Algeria","Bahrain","Belarus","Egypt","Ghana","Iran","Ivory Coast","Jordan","Kazakhstan","Kenya","Kuwait","Lebanon","Morocco","Nigeria","Oman","Pakistan","Qatar","Russia","Saudi Arabia","South Africa","Tunisia","Turkey","Ukraine","United Arab Emirates","Indonesia","India","Malaysia","Philippines","Sri Lanka","Thailand","Vietnam","Argentina","Brazil","Chile","Colombia","Costa Rica","Dominican Republic","Ecuador","El Salvador","Guatemala","Honduras","Jamaica","Mexico","Nicaragua","Panama","Peru","Trinidad and Tobago","Uruguay","Estonia","Germany","Greece","Guatemala","Honduras","Hungary","Italy","Jamaica","Jordan","Kenya","Latvia","Nigeria","Panama","Serbia","Slovakia","Slovenia","Taiwan","Trinidad and Tobago","Turkey"]

    def to_data_frame(self):

        def addGlobalRole(userid, therapyAreas, df_roleSetup, df_countries, df_therapyAreas):
            template = self.df_template
            for i, therapyArea in enumerate(therapyAreas):
                if i == 0:
                    tmp_df = df_roleSetup[df_roleSetup["Persona"] == role]
                    tmp_df = tmp_df.drop(columns=['Persona'])

                if i > 0:
                    tmp_df = df_roleSetup[df_roleSetup["Persona"] == role]
                    tmp_df = tmp_df.drop(columns=['Persona'])
                    tmp_df = tmp_df.dropna(subset=['franchise__c'])
                    tmp_df = tmp_df[tmp_df['country__c'] != "Global"]

                id_country = df_countries.loc[df_countries["Country Name"] == market]
                id_global = df_countries.loc[df_countries["Country Name"] == "Global"]
                id_therapyArea = df_therapyAreas.loc[df_therapyAreas["Picklist Value Label"] == therapyAreas[i]]

                id_country_val = id_country.iloc[0, 4]
                id_global_val = id_global.iloc[0, 4]
                id_therapy_area_val = id_therapyArea.iloc[0, 1]

                tmp_df = tmp_df.replace({"{LOC}": id_country_val,
                                         "Global": id_global_val,
                                         "user_id": userid,
                                         "{Therapy Area}": id_therapy_area_val})
                template = tmp_df.append(template)
            return template

        def addLocalRole(iterator, userid, df_roleSetup, df_countries, df_specialCountries):
            if iterator == 0:
                tmp_df = df_roleSetup[df_roleSetup["Persona"] == role]
                tmp_df = tmp_df.drop(columns=["Persona"])

            if iterator > 0:
                tmp_df = df_roleSetup[df_roleSetup["Persona"] == role]
                tmp_df = tmp_df.drop(columns=['Persona'])
                tmp_df = tmp_df.dropna(subset=['country__c'])
                tmp_df = tmp_df[tmp_df['country__c'] != "Global"]

            id_country = df_countries.loc[df_countries["Country Name"] == market]
            id_global = df_countries.loc[df_countries["Country Name"] == "Global"]

            id_country_val = id_country.iloc[0, 4]
            id_global_val = id_global.iloc[0, 4]

            # ADDING ROW(S) FOR THE 'SUBMISSION COORDINATOR' ROLE (BASED ON 'SPECIAL COUNTRY ROLES' LIST)
            if market in self.specialCountryRoles:
                df_curr_specialCountries = df_specialCountries.loc[(df_specialCountries["COUNTRY_NAME"] == market) & (df_specialCountries["ROLE_NAME"] == role)]
                df_curr_specialCountries = df_curr_specialCountries.drop(columns=["COUNTRY_NAME", "ROLE_NAME"])
                tmp_df = pd.concat([tmp_df, df_curr_specialCountries])  # ZLACZENIE WYCINKA MATRYCY Z WIERSZEM "SUBMISSION COORDINATOR"

            tmp_df = tmp_df.replace({"{LOC}": id_country_val,
                                     "Global": id_global_val,
                                     "user_id": userid})

            return tmp_df

        input = self.input
        template = self.df_template

        for role in input.get("roles"):
            for iterator, market in enumerate(input.get("markets")):
                if role in self.therapyAreasRoles:
                    tmp_df = addGlobalRole(userid=input.get("MUDID"),
                                           therapyAreas=input.get("therapyAreas"),
                                           df_roleSetup=self.df_roleSetup,
                                           df_countries=self.df_countries,
                                           df_therapyAreas=self.df_therapyAreas)
                    template = template.append(tmp_df)

                if role not in self.therapyAreasRoles:
                    tmp_df = addLocalRole(iterator=iterator,
                                          userid=input.get("MUDID"),
                                          df_roleSetup=self.df_roleSetup,
                                          df_countries=self.df_countries,
                                          df_specialCountries=self.df_specialRoles)
                    template = template.append(tmp_df)

        return template


class ExternalUserRoleSetup(File):
    template_name = "fresh_template_user_role_setup.xlsx"

    def __init__(self, input):
        super().__init__()
        self.input = input
        self.df_template = pd.read_excel(self.template_name)
        self.df_roleSetup = pd.read_excel(self.matrice_name, sheet_name="ROLE_SETUP")
        self.df_products = pd.read_excel(self.matrice_name, sheet_name="Products")
        self.productRoles = ['Global Commercial Agency', 'Global Medical Agency']

    def to_data_frame(self):

        def addLocalRole(iterator, userid, role, agency, market, df_roleSetup, df_countries):
            if iterator == 0:
                tmp_df = df_roleSetup[df_roleSetup["Persona"] == role]
                tmp_df = tmp_df.drop(columns=["Persona"])

            if iterator > 0:
                tmp_df = df_roleSetup[df_roleSetup["Persona"] == role]
                tmp_df = tmp_df.drop(columns=['Persona'])
                tmp_df = tmp_df.dropna(subset=['country__c'])
                tmp_df = tmp_df[tmp_df['country__c'] != "Global"]

            id_country = df_countries.loc[df_countries["Country Name"] == market]
            id_global = df_countries.loc[df_countries["Country Name"] == "Global"]

            id_country_val = id_country.iloc[0, 4]
            id_global_val = id_global.iloc[0, 4]

            tmp_df = tmp_df.replace({"{LOC}": id_country_val,
                                     "Global": id_global_val,
                                     "user_id": userid,
                                     "{Creative Agency}": agency})

            return tmp_df

        def addGlobalRole(iterator, userid, role, agency, market, product, df_roleSetup, df_countries, df_products):
            if iterator == 0:
                tmp_df = df_roleSetup[df_roleSetup["Persona"] == role]
                tmp_df = tmp_df.drop(columns=["Persona"])

            if iterator > 0:
                tmp_df = df_roleSetup[df_roleSetup["Persona"] == role]
                tmp_df = tmp_df.drop(columns=['Persona'])
                tmp_df = tmp_df.dropna(subset=['product__c'])
                tmp_df = tmp_df[tmp_df['country__c'] != "Global"]

            id_country = df_countries.loc[df_countries["Country Name"] == market]
            id_global = df_countries.loc[df_countries["Country Name"] == "Global"]
            id_product = df_products.loc[df_products["Product Name"] == product]

            id_country_val = id_country.iloc[0, 4]
            id_global_val = id_global.iloc[0, 4]
            id_product_val = id_product.iloc[0, 1]

            tmp_df = tmp_df.replace({"{LOC}": id_country_val,
                                     "Global": id_global_val,
                                     "user_id": userid,
                                     "{Creative Agency}": agency,
                                     "{Product}": id_product_val})

            return tmp_df

        input = self.input
        template = self.df_template

        for role in input.get("roles"):
            if role not in self.productRoles:
                for iterator, market in enumerate(input.get("markets")):
                    tmp_df = addLocalRole(iterator=iterator,
                                          userid=input.get("MUDID"),
                                          role=role,
                                          agency=input.get("agency"),
                                          market=market,
                                          df_roleSetup=self.df_roleSetup,
                                          df_countries=self.file.df_countries)
                    template = template.append(tmp_df)

            if role in self.productRoles:
                for iterator, product in enumerate(input.get("products")):
                    tmp_df = addGlobalRole(iterator=iterator,
                                           userid=input.get("MUDID"),
                                           role=role,
                                           agency=input.get("agency"),
                                           market="Global",
                                           product=product,
                                           df_roleSetup=self.df_roleSetup,
                                           df_countries=self.df_countries,
                                           df_products=self.df_products)
                    template = template.append(tmp_df)

        return template


class UsUsersUserRoleSetup(File):
    template_name = "fresh_template_user_role_setup.xlsx"

    def __init__(self, input):
        super().__init__()
        self.df_template = pd.read_excel(self.template_name)
        self.df_roleSetup = pd.read_excel(self.matrice_name, sheet_name="US_ROLE_SETUP")
        self.input = input

    def to_data_frame(self):

        def add_role(iterator, userid, role, agencyName, df_roleSetup):
            if iterator == 0:
                tmp_df = df_roleSetup[df_roleSetup["Persona"] == role]
                tmp_df = tmp_df.drop(columns=["Persona"])

            if iterator > 0:
                tmp_df = df_roleSetup[df_roleSetup["Persona"] == role]
                tmp_df = tmp_df.drop(columns=['Persona'])
                tmp_df = tmp_df.dropna(subset=['country__c'])
                tmp_df = tmp_df[tmp_df['country__c'] != "Global"]


            tmp_df = tmp_df.replace({"user_id": userid,
                                     "agency__c": agencyName})
            return tmp_df

        for iterator, role in enumerate(self.input.get("roles")):
            tmp_df = add_role(iterator      = iterator,
                             userid         = self.input.get("MUDID"),
                             agencyName     = self.input.get('agency'),
                             df_roleSetup   = self.df_roleSetup,
                             role           = role)
            self.df_template = self.df_template.append(tmp_df)

        return self.df_template


class Page(tk.Frame):
    def __init__(self, *args, **kwargs):
        tk.Frame.__init__(self, *args, **kwargs)

    def show(self):
        self.lift()


class InternalUsers(Page):
    def __init__(self, height, width):
        Page.__init__(self, bg='#ffab45')

        self.labelDict = {}
        self.buttonList = []
        self.listboxList = []
        self.entryList = []

        self.roles = ['Local Commercial Content Owner', 'Local Commercial Reviewer', 'Local Commercial Approver',
                      'Commercial LOC Custodian',
                      'Local Medical Content Owner', 'Local Medical Reviewer', 'Local Medical Approver',
                      'Medical LOC Custodian',
                      'Local Pharmacovigilance Reviewer', 'Local Pharmacovigilance Approver', 'Local Legal Reviewer',
                      'Local Legal Approver',
                      'Local Regulatory Reviewer', 'Local Regulatory Approver', 'Local Other Reviewer/Approver',
                      'Local Commercial Submission Coordinator',
                      'Local Medical Submission Coordinator', 'Local Commercial Coordinator',
                      'Local Medical Coordinator', 'Local Commercial Agency',
                      'Local Medical Agency', 'Local Read Only', 'Knowledge Center', 'Local Custodian',
                      'Service Provider', 'Librarian', 'GEM-MEM-MOC', 'DPM-MOC',
                      'MOC-Production', 'DPM-DC', 'Document Admin', 'Global Commercial Content Owner',
                      'Global Commercial Reviewer', 'Global Commercial Approver',
                      'Global Commercial Brand Portal Manager',
                      'Global Commercial Brand Portal Admin', 'Commercial Global Custodian',
                      'Global Medical Content Owner', 'Global Medical Reviewer',
                      'Global Medical Approver', 'Global Medical Brand Portal Manager',
                      'Global Medical Brand Portal Admin', 'Medical Global Custodian', 'Global Other Reviewer/Approver',
                      'Global Commercial Agency', 'Global Medical Agency',
                      'Global Read Only', 'Global Custodian', 'Librarian Global Portal Manager',
                      'Librarian Global Portal Admin']
        self.markets = ['Albania', 'Algeria', 'Antigua and Barbuda', 'Argentina', 'Armenia', 'Aruba', 'Australia',
                        'Austria',
                        'Azerbaijan', 'Bahamas', 'Bahrain', 'Bangladesh', 'Barbados', 'Belarus', 'Belgium', 'Benin',
                        'Bermuda',
                        'Bolivia', 'Bosnia and Herzegovina', 'Botswana', 'Brazil', 'British Virgin Islands', 'Brunei',
                        'Bulgaria',
                        'Cambodia', 'Cameroon', 'Canada', 'Cayman Islands', 'Chile', 'China', 'Colombia', 'Costa Rica',
                        'Croatia', 'Curacao',
                        'Cyprus', 'Czech Republic', 'Democratic Republic of the Congo', 'Denmark', 'Dominica',
                        'Dominican Republic', 'Ecuador',
                        'Egypt', 'El Salvador', 'Estonia', 'Ethiopia', 'Finland', 'France', 'Gabon', 'Georgia',
                        'Germany', 'Ghana', 'Global',
                        'Greece', 'Grenada', 'Guatemala', 'Guinea', 'Guyana', 'Haiti', 'Honduras', 'Hong Kong',
                        'Hungary', 'Iceland', 'India',
                        'Indonesia', 'Iran', 'Ireland', 'Israel', 'Italy', 'Ivory Coast', 'Jamaica', 'Japan', 'Jordan',
                        'Kazakhstan', 'Kenya',
                        'Korea', 'Kuwait', 'Laos', 'Latvia', 'Lebanon', 'Lithuania', 'Luxembourg', 'Macau', 'Macedonia',
                        'Madagascar', 'Malawi',
                        'Malaysia', 'Malta', 'Mauritius', 'Mexico', 'Moldova', 'Montenegro', 'Morocco', 'Mozambique',
                        'Myanmar', 'Namibia',
                        'Netherlands', 'Netherlands Antilles', 'New Zealand', 'Nicaragua', 'Nigeria', 'Nordic Cluster',
                        'Norway', 'Oman',
                        'Pakistan', 'Panama', 'Papua New Guinea', 'Paraguay', 'Peru', 'Philippines', 'Poland',
                        'Portugal', 'Puerto Rico',
                        'Qatar', 'Region / Cluster / Hub', 'Republic of the Congo', 'Romania', 'Russia', 'Rwanda',
                        'Saint Kitts and Nevis',
                        'Saint Lucia', 'Saint Martin', 'Saint Vincent and Grenadines', 'Saudi Arabia', 'Senegal',
                        'Serbia', 'Seychelles',
                        'Singapore', 'Slovakia', 'Slovenia', 'South Africa', 'Spain', 'Sri Lanka', 'Suriname', 'Sweden',
                        'Switzerland', 'Taiwan',
                        'Tanzania', 'Thailand', 'Togo', 'Trinidad and Tobago', 'Tunisia', 'Turkey', 'Uganda', 'Ukraine',
                        'United Arab Emirates',
                        'United Kingdom', 'United States', 'Uruguay', 'Uzbekistan', 'Venezuela', 'Vietnam', 'Zambia']

        matriceSourceTherapyAreas = pd.read_excel('matrice_source.xlsx', sheet_name='THERAPY_AREA_SETUP')

        self.securityProfiles = ['Commercial Content Owner', 'Medical Content Owner', 'PromoMats User',
                                 'Read-Only PromoMats User', 'Librarian', 'LOC Custodian', 'Knowledge Center',
                                 'GEM/MEM', 'Document Admin', 'Commercial Content Owner + Brand Portal Admin',
                                 'Medical Content Owner + Brand Portal Admin', 'Librarian + Brand Portal Admin']
        self.therapyAreas = ['Classic and Established Medicines', 'Immunology & Specialty Medicines (I&SM)', 'Oncology',
                             'Specialty and Primary Care', 'Respiratory', 'Vaccines', 'ViiV']

        self.fileTypeSelected = None
        self.selectedRoles = None
        self.selectedMarkets = None
        self.selectedSecurityProfiles = None
        self.MUDID = None
        self.UserName = None
        self.selectedTherapyAreas = None

        listbox_names = ["RoleList", "MarketList", "SecurityProfilesList", "therapyAreasList"]
        entry_names = ["MUDID", "userName"]

        # MAIN PART
        # Random colour generation
        r = lambda: random.randint(0, 255)
        bg = '#%02X%02X%02X' % (r(), r(), r())
        frameMain = tk.Frame(self, bg=bg)         # #ffab45 <-- oryginal orange colour
        frameMain.place(relx=0.00, rely=0.00, relwidth=1, relheight=1)

        self.addLabel(text="Which file do you want to create?", where_add=frameMain, relx=0.1, rely=0.015, relwidth=0.4,
                      relheight=0.05)
        self.addLabel(text="Please select role(s)", where_add=frameMain, relx=0.1, rely=0.33)
        self.addLabel(text="Please select market(s)",  where_add=frameMain, relx=0.55, rely=0.33)
        self.addLabel(text="Please select security profile",  where_add=frameMain, relx=0.1, rely=0.6)
        self.addLabel(text="User ID", where_add=frameMain, relx=0.45, rely=0.7)
        self.addLabel(text="User Name",  where_add=frameMain, relx=0.1, rely=0.7)
        self.addLabel(text="Please select therapy area(s)",  where_add=frameMain, relx=0.55, rely=0.6)
        self.addButton(text="Content Plan", where_add=frameMain, relx=0.55, rely=0.015, relwidth=0.15, relheight=0.05,
                       command=lambda: fileTypeSelected("Content Plan"))
        self.addButton(text="User Role Setup", where_add=frameMain, relx=0.72, rely=0.015, relwidth=0.15,
                       relheight=0.05, command=lambda: fileTypeSelected("User Role Setup"))
        self.addButton(text="Select", where_add=frameMain, relx=0.36, rely=0.33, relwidth=0.15, relheight=0.05,
                       command=lambda: selectedRoles(self.listboxList, listbox_names[0]))
        self.addButton(text="Select", where_add=frameMain, relx=0.81, rely=0.33, relwidth=0.15, relheight=0.05,
                       command=lambda: selectedMarkets(self.listboxList, listbox_names[1]))
        self.addButton(text="Select", where_add=frameMain, relx=0.36, rely=0.6, relwidth=0.15, relheight=0.05,
                       command=lambda: selectedSecurityProfiles(self.listboxList, listbox_names[2]))
        self.addButton(text="Select", where_add=frameMain, relx=0.45, rely=0.735, relwidth=0.15, relheight=0.05,
                       command=lambda: selectedMUDID(self.entryList, entry_names[0]))
        self.addButton(text="Select", where_add=frameMain, relx=0.1, rely=0.735, relwidth=0.15, relheight=0.05,
                       command=lambda: selectedUserName(self.entryList, entry_names[1]))
        self.addButton(text="Select", where_add=frameMain, relx=0.81, rely=0.6, relwidth=0.15, relheight=0.05,
                       command=lambda: selectedTherapyAreas(self.listboxList, listbox_names[3]))
        self.addButton(text="Create csv file", where_add=frameMain, relx=0.8, rely=0.75, relwidth=0.15, relheight=0.05,
                       command=lambda: generateFile())
        self.addListBox(where_add=frameMain, select_mode="multiple", width=40, relx=0.1, rely=0.125,
                        contents=self.roles, name="RoleList")
        self.addListBox(where_add=frameMain, select_mode="multiple", width=40, relx=0.55, rely=0.125,
                        contents=self.markets, name="MarketList")
        self.addListBox(where_add=frameMain, select_mode="multiple", width=40, relx=0.1, rely=0.4,
                        contents=self.securityProfiles, name="SecurityProfilesList")
        self.addListBox(where_add=frameMain, select_mode="multiple", width=40, relx=0.55, rely=0.4,
                        contents=self.therapyAreas, name="therapyAreasList")
        self.addEntry(where_add=frameMain, relx=0.53, rely=0.70, name="MUDID")
        self.addEntry(where_add=frameMain, relx=0.215, rely=0.70, name="userName")

        # SUMMARY
        frame_summary = tk.Frame(self, bg='#c7b6a1')
        frame_summary.place(relx=0.00, rely=0.80, relwidth=1, relheight=0.15)

        self.addLabel(text="SUMMARY",  where_add=frame_summary, relx=0.05, rely=0.10, relwidth=0.15, relheight=0.1)
        self.addLabel(text="File Type: ",  where_add=frame_summary, relx=0.05, rely=0.2, relwidth=0.20, relheight=0.1)
        self.addLabel(text="",  where_add=frame_summary, relx=0.25, rely=0.2, relwidth=0.7, relheight=0.10)                     #9
        self.addLabel(text="Roles: ",  where_add=frame_summary, relx=0.05, rely=0.30, relwidth=0.20, relheight=0.1)
        self.addLabel(text="",  where_add=frame_summary, relx=0.25, rely=0.30, relwidth=0.7, relheight=0.1)                     #11
        self.addLabel(text="Markets: ",  where_add=frame_summary, relx=0.05, rely=0.4, relwidth=0.20, relheight=0.1)
        self.addLabel(text="",  where_add=frame_summary, relx=0.25, rely=0.4, relwidth=0.7, relheight=0.1)                      #13
        self.addLabel(text="Security profile: ",  where_add=frame_summary, relx=0.05, rely=0.5, relwidth=0.20, relheight=0.1)
        self.addLabel(text="",  where_add=frame_summary, relx=0.25, rely=0.5, relwidth=0.7, relheight=0.1)                      #15
        self.addLabel(text="User ID: ",  where_add=frame_summary, relx=0.05, rely=0.6, relwidth=0.20, relheight=0.1)
        self.addLabel(text="",  where_add=frame_summary, relx=0.25, rely=0.6, relwidth=0.7, relheight=0.1)                      #17
        self.addLabel(text="User Name: ",  where_add=frame_summary, relx=0.05, rely=0.7, relwidth=0.20, relheight=0.1)
        self.addLabel(text="",  where_add=frame_summary, relx=0.25, rely=0.7, relwidth=0.7, relheight=0.1)                      #19
        self.addLabel(text="Therapy Areas: ",  where_add=frame_summary, relx=0.05, rely=0.8, relwidth=0.20, relheight=0.1)
        self.addLabel(text="", where_add=frame_summary, relx=0.25, rely=0.8, relwidth=0.7, relheight=0.1)                       #21


        #FUNCTIONS HANDLING BUTTONS
        def fileTypeSelected(filetype):
            self.fileTypeSelected = filetype
            self.labelDict.get(9)['text'] = filetype
            print(filetype)

        def selectedRoles(listBoxDictList, name):
            for element in listBoxDictList:
                if element.get("Name") == name:
                    list = element.get("ListBox object")
                    self.selectedRoles = [list.get(idx) for idx in list.curselection()]
                    print(self.selectedRoles)
            self.labelDict.get(11)['text'] = self.selectedRoles

        def selectedMarkets(listBoxDictList, name):
            for element in listBoxDictList:
                if element.get("Name") == name:
                    list = element.get("ListBox object")
                    self.selectedMarkets = [list.get(idx) for idx in list.curselection()]
                    print(self.selectedMarkets)
            self.labelDict.get(13)['text'] = self.selectedMarkets

        def selectedSecurityProfiles(listBoxDictList, name):
            for element in listBoxDictList:
                if element.get("Name") == name:
                    list = element.get("ListBox object")
                    self.selectedSecurityProfiles = [list.get(idx) for idx in list.curselection()]
                    print(self.selectedSecurityProfiles)
            self.labelDict.get(15)['text'] = self.selectedSecurityProfiles

        def selectedMUDID(entryDictList, name):
            for element in entryDictList:
                if element.get("Name") == name:
                    entry = element.get("Entry object")
                    self.MUDID = entry.get()
                    print(self.MUDID)
            self.labelDict.get(17)['text'] = self.MUDID

        def selectedUserName(entryDictList, name):
            for element in entryDictList:
                if element.get("Name") == name:
                    entry = element.get("Entry object")
                    self.UserName = entry.get()
                    print(self.UserName)
            self.labelDict.get(19)['text'] = self.UserName

        def selectedTherapyAreas(therapyAreasDictList, name):
            for element in therapyAreasDictList:
                if element.get("Name") == name:
                    list = element.get("ListBox object")
                    self.selectedTherapyAreas = [list.get(idx) for idx in list.curselection()]
                    print(self.selectedTherapyAreas)
            self.labelDict.get(21)['text'] = self.selectedTherapyAreas

        def generateFile():

            user_input = {"fileType": self.fileTypeSelected,
                          "roles": self.selectedRoles,
                          "markets": self.selectedMarkets,
                          "MUDID": self.MUDID,
                          "securityProfile": self.selectedSecurityProfiles,
                          "userName": self.UserName,
                          "therapyAreas": self.selectedTherapyAreas}

            if user_input.get("fileType") == "Content Plan":
                contentPlan = ContentPlan(user_input).to_data_frame()
                pathfile = '_1_{}_content_plan.csv'.format(user_input.get("userName"))
                contentPlan.to_csv(pathfile, index=False, header=True)

            if user_input.get("fileType") == "User Role Setup":
                userRoleSetup = InternalUserRoleSetup(user_input).to_data_frame()
                pathfile = '_2_{}_user_role_setup.csv'.format(user_input.get("userName"))
                userRoleSetup.to_csv(pathfile, index=False, header=True)


    # FUNCTIONS ADDING WIDGETS
    def addLabel(self, text, where_add, relx, rely, relwidth=None, relheight=None):
        label = tk.Label(where_add, text=text)
        if relwidth and relheight is not None:
            label.place(relx=relx, rely=rely, relwidth=relwidth, relheight=relheight)
        else:
            label.place(relx=relx, rely=rely)
        dict_to_add = {len(self.labelDict): label}
        self.labelDict.update(dict_to_add)

    def addButton(self, text, where_add, relx, rely, relwidth, relheight, command=None):
        button = tk.Button(where_add, text=text)
        button.place(relx=relx, rely=rely, relwidth=relwidth, relheight=relheight)
        if command is not None:
            button.configure(command=command)
        self.buttonList.append(button)

    def addListBox(self, where_add, select_mode, width, relx, rely, contents, name):
        ListBox = tk.Listbox(where_add, selectmode=select_mode, width=width)
        ListBox.place(relx=relx, rely=rely)
        for index, each_item in enumerate(contents):
            ListBox.insert(index, each_item)
        dict_to_add = {"Name": name,
                       "ListBox object": ListBox}
        self.listboxList.append(dict_to_add)

    def addEntry(self, where_add, relx, rely, name):
        Entry = tk.Entry(where_add)
        Entry.place(relx=relx, rely=rely)
        dict_to_add = {"Name": name,
                       "Entry object": Entry}
        self.entryList.append(dict_to_add)


class ExternalUsers(Page):
    def __init__(self, *args, **kwargs):
        Page.__init__(self, bg='#ffab45')

        matriceSourceProducts = pd.read_excel('matrice_source.xlsx', sheet_name='Products')
        matriceSourceCountries = pd.read_excel('matrice_source.xlsx', sheet_name='Countries')

        self.labelDict = {}
        self.buttonList = []
        self.listboxList = []
        self.entryList = []

        self.roles = ['Local Commercial Agency', 'Local Medical Agency', 'Global Commercial Agency', 'Global Medical Agency']
        self.markets = matriceSourceCountries['Country Name'].tolist()
        self.products = matriceSourceProducts['Product Name'].tolist()

        self.fileTypeSelected = None
        self.selectedRoles = None
        self.selectedMarkets = None
        self.MUDID = None
        self.selectedProducts = []
        self.UserName = None
        self.agencyName = None

        listbox_names = ["RoleList", "MarketList", "ProductsList", "therapyAreasList"]
        entry_names = ["MUDID", "userName", "agencyName"]

        # MAIN PART
        # Random colour generation
        r = lambda: random.randint(0, 255)
        bg = '#%02X%02X%02X' % (r(), r(), r())
        frameMain = tk.Frame(self, bg='#ffab45')  # #ffab45 <-- oryginal orange colour
        frameMain.place(relx=0.00, rely=0.00, relwidth=1, relheight=1)

        self.addLabel(text="Please select role(s)", where_add=frameMain, relx=0.1, rely=0.33)
        self.addLabel(text="Please select market(s)", where_add=frameMain, relx=0.55, rely=0.33)
        self.addLabel(text="Please select products", where_add=frameMain, relx=0.1, rely=0.605)
        self.addLabel(text="User ID", where_add=frameMain, relx=0.55, rely=0.525)
        self.addLabel(text="User Name", where_add=frameMain, relx=0.55, rely=0.475)
        self.addLabel(text="Agency Name", where_add=frameMain, relx=0.55, rely=0.575)
        self.addButton(text="Select", where_add=frameMain, relx=0.36, rely=0.33, relwidth=0.15, relheight=0.05,
                       command=lambda: selectedRoles(self.listboxList, listbox_names[0]))
        self.addButton(text="Select", where_add=frameMain, relx=0.81, rely=0.33, relwidth=0.15, relheight=0.05,
                       command=lambda: selectedMarkets(self.listboxList, listbox_names[1]))
        self.addButton(text="Select", where_add=frameMain, relx=0.36, rely=0.6, relwidth=0.15, relheight=0.05,
                       command=lambda: selectedProduct(self.listboxList, listbox_names[2]))
        self.addButton(text="Select", where_add=frameMain, relx=0.92, rely=0.525, relwidth=0.07, relheight=0.025,
                       command=lambda: selectedMUDID(self.entryList, entry_names[0]))
        self.addButton(text="Select", where_add=frameMain, relx=0.92, rely=0.475, relwidth=0.07, relheight=0.025,
                       command=lambda: selectedUserName(self.entryList, entry_names[1]))
        self.addButton(text="Create User Role Setup", where_add=frameMain, relx=0.75, rely=0.70, relwidth=0.2, relheight=0.05,
                       command=lambda: generateFile())
        self.addButton(text="Select", where_add=frameMain, relx=0.92, rely=0.575, relwidth=0.07, relheight=0.025,
                       command=lambda: selectedAgency(self.entryList, entry_names[2]))
        self.addButton(text="Clear selection", where_add=frameMain, relx=0.36, rely=0.65, relwidth=0.15, relheight=0.05,
                       command=lambda: clearProducts())
        self.addListBox(where_add=frameMain, select_mode="multiple", width=40, relx=0.1, rely=0.125,
                        contents=self.roles, name="RoleList")
        self.addListBox(where_add=frameMain, select_mode="multiple", width=40, relx=0.55, rely=0.125,
                        contents=self.markets, name="MarketList")
        self.addListBox(where_add=frameMain, select_mode="multiple", width=40, relx=0.1, rely=0.4,
                        contents=self.products, name="ProductsList")

        self.addEntry(where_add=frameMain, relx=0.7, rely=0.525, name="MUDID")
        self.addEntry(where_add=frameMain, relx=0.7, rely=0.475, name="userName")
        self.addEntry(where_add=frameMain, relx=0.7, rely=0.575, name="agencyName")
        self.addEntry(where_add=frameMain, relx=0.1, rely=0.625, name="products")

        def updateListbox(data, listbox):
            # Clear the listbox
            listbox.delete(0, 'end')

            # Add items to listbox
            for item in data:
                listbox.insert('end', item)

        def checkProductListbox(event):
            """Create a function to check entry vs listbox"""
            # grab what was typed
            entry = self.entryList[3].get("Entry object")
            typed = entry.get()

            # if nothing was typed reset the listbox
            if typed == '':
                data = self.products
            else:
                data = []
                for item in self.products:
                    if typed.lower() in item.lower():
                        data.append(item)

            # update our listbox with selected items
            productListbox = self.listboxList[2].get("ListBox object")
            updateListbox(data, productListbox)

        self.entryList[3].get("Entry object").bind("<KeyRelease>", checkProductListbox)



        # SUMMARY
        frame_summary = tk.Frame(self, bg='#c7b6a1')
        frame_summary.place(relx=0.00, rely=0.80, relwidth=1, relheight=0.15)

        self.addLabel(text="SUMMARY",  where_add=frame_summary, relx=0.05, rely=0.10, relwidth=0.15, relheight=0.1)
        self.addLabel(text="Roles: ",  where_add=frame_summary, relx=0.05, rely=0.20, relwidth=0.20, relheight=0.1)
        self.addLabel(text="",  where_add=frame_summary, relx=0.25, rely=0.20, relwidth=0.7, relheight=0.1)                     #8
        self.addLabel(text="Markets: ",  where_add=frame_summary, relx=0.05, rely=0.3, relwidth=0.20, relheight=0.1)
        self.addLabel(text="",  where_add=frame_summary, relx=0.25, rely=0.3, relwidth=0.7, relheight=0.1)                      #10
        self.addLabel(text="User ID: ",  where_add=frame_summary, relx=0.05, rely=0.4, relwidth=0.20, relheight=0.1)
        self.addLabel(text="",  where_add=frame_summary, relx=0.25, rely=0.4, relwidth=0.7, relheight=0.1)                      #12
        self.addLabel(text="User Name: ",  where_add=frame_summary, relx=0.05, rely=0.5, relwidth=0.20, relheight=0.1)
        self.addLabel(text="",  where_add=frame_summary, relx=0.25, rely=0.5, relwidth=0.7, relheight=0.1)                      #14
        self.addLabel(text="Products: ",  where_add=frame_summary, relx=0.05, rely=0.6, relwidth=0.20, relheight=0.1)
        self.addLabel(text="", where_add=frame_summary, relx=0.25, rely=0.6, relwidth=0.7, relheight=0.1)                       #16
        self.addLabel(text="Agency: ",  where_add=frame_summary, relx=0.05, rely=0.7, relwidth=0.20, relheight=0.1)
        self.addLabel(text="", where_add=frame_summary, relx=0.25, rely=0.7, relwidth=0.7, relheight=0.1)                       #18
        #print(self.labelDict)

        def selectedRoles(listBoxDictList, name):
            for element in listBoxDictList:
                if element.get("Name") == name:
                    list = element.get("ListBox object")
                    self.selectedRoles = [list.get(idx) for idx in list.curselection()]
                    print(self.selectedRoles)
            self.labelDict.get(8)['text'] = self.selectedRoles

        def selectedMarkets(listBoxDictList, name):
            for element in listBoxDictList:
                if element.get("Name") == name:
                    list = element.get("ListBox object")
                    self.selectedMarkets = [list.get(idx) for idx in list.curselection()]
                    print(self.selectedMarkets)
            self.labelDict.get(10)['text'] = self.selectedMarkets

        def selectedProduct(listBoxDictList, name):
            for element in listBoxDictList:
                if element.get("Name") == name:
                    list = element.get("ListBox object")
                    selection = [list.get(idx) for idx in list.curselection()]

            for i in range(len(selection)):
                self.selectedProducts.append(selection[i])

            self.labelDict.get(16)['text'] = self.selectedProducts
            print(self.selectedProducts)

        def clearProducts():
            self.selectedProducts = []
            self.labelDict.get(16)['text'] = self.selectedProducts
            print(self.selectedProducts)

        def selectedMUDID(entryDictList, name):
            for element in entryDictList:
                if element.get("Name") == name:
                    entry = element.get("Entry object")
                    self.MUDID = entry.get()
                    print(self.MUDID)
            self.labelDict.get(12)['text'] = self.MUDID

        def selectedUserName(entryDictList, name):
            for element in entryDictList:
                if element.get("Name") == name:
                    entry = element.get("Entry object")
                    self.UserName = entry.get()
                    print(self.UserName)
            self.labelDict.get(14)['text'] = self.UserName

        def selectedAgency(entryDictList, name):
            for element in entryDictList:
                if element.get("Name") == name:
                    entry = element.get("Entry object")
                    self.agencyName = entry.get()
                    print(self.agencyName)
            self.labelDict.get(18)['text'] = self.agencyName

        def generateFile():

            user_input = {"roles": self.selectedRoles,
                          "markets": self.selectedMarkets,
                          "MUDID": self.MUDID,
                          "products": self.selectedProducts,
                          "userName": self.UserName,
                          'agency': self.agencyName}

            print(user_input)
            userRoleSetup = ExternalUserRoleSetup(user_input).to_data_frame()
            pathfile = '_2_{}_EXTERNAL_user_role_setup.csv'.format(user_input.get("userName"))
            userRoleSetup.to_csv(pathfile, index=False, header=True)

    # FUNCTIONS ADDING WIDGETS
    def addLabel(self, text, where_add, relx, rely, relwidth=None, relheight=None):
        label = tk.Label(where_add, text=text)
        if relwidth and relheight is not None:
            label.place(relx=relx, rely=rely, relwidth=relwidth, relheight=relheight)
        else:
            label.place(relx=relx, rely=rely)
        dict_to_add = {len(self.labelDict): label}
        self.labelDict.update(dict_to_add)

    def addButton(self, text, where_add, relx, rely, relwidth, relheight, command=None):
        button = tk.Button(where_add, text=text)
        button.place(relx=relx, rely=rely, relwidth=relwidth, relheight=relheight)
        if command is not None:
            button.configure(command=command)
        self.buttonList.append(button)

    def addListBox(self, where_add, select_mode, width, relx, rely, contents, name):
        ListBox = tk.Listbox(where_add, selectmode=select_mode, width=width)
        ListBox.place(relx=relx, rely=rely)
        for index, each_item in enumerate(contents):
            ListBox.insert(index, each_item)
        dict_to_add = {"Name": name,
                       "ListBox object": ListBox}
        self.listboxList.append(dict_to_add)

    def addEntry(self, where_add, relx, rely, name):
        Entry = tk.Entry(where_add)
        Entry.place(relx=relx, rely=rely)
        dict_to_add = {"Name": name,
                       "Entry object": Entry}
        self.entryList.append(dict_to_add)


class UsUsers(Page):

    def __init__(self, *args, **kwargs):

        Page.__init__(self, bg='#ffab45')

        matriceSourceRoles = pd.read_excel('matrice_source.xlsx', sheet_name='US_ROLE_SETUP')

        self.labelDict = {}
        self.buttonList = []
        self.listboxList = []
        self.entryList = []

        self.roles = unique_values(matriceSourceRoles['Persona'].tolist())
        self.securityProfiles = ['Commercial Content Owner', 'Medical Content Owner', 'PromoMats User',
                                 'Read-Only PromoMats User', 'Librarian', 'LOC Custodian', 'Knowledge Center',
                                 'GEM/MEM', 'Document Admin', 'Commercial Content Owner + Brand Portal Admin',
                                 'Medical Content Owner + Brand Portal Admin', 'Librarian + Brand Portal Admin']
        self.fileTypeSelected = None
        self.selectedRoles = None
        self.selectedMarkets = None
        self.MUDID = None
        self.selectedProducts = []
        self.UserName = None
        self.agencyName = None
        self.selectedSecurityProfiles = None

        listbox_names = ["RoleList", "MarketList", "ProductsList", "therapyAreasList", "SecurityProfilesList"]
        entry_names = ["MUDID", "userName", "agencyName"]

        # MAIN PART
        # Random colour generation
        r = lambda: random.randint(0, 255)
        bg = '#%02X%02X%02X' % (r(), r(), r())
        frameMain = tk.Frame(self, bg='#ffab45')  # #ffab45 <-- oryginal orange colour
        frameMain.place(relx=0.00, rely=0.00, relwidth=1, relheight=1)

        self.addLabel(text="Which file do you want to create?", where_add=frameMain, relx=0.1, rely=0.015, relwidth=0.4,
                      relheight=0.05)
        self.addLabel(text="Please select role(s)", where_add=frameMain, relx=0.1, rely=0.33)
        self.addLabel(text="User ID", where_add=frameMain, relx=0.55, rely=0.525)
        self.addLabel(text="User Name", where_add=frameMain, relx=0.55, rely=0.475)
        self.addLabel(text="Agency Name", where_add=frameMain, relx=0.55, rely=0.575)
        self.addLabel(text="Please select security profile",  where_add=frameMain, relx=0.1, rely=0.6)
        self.addButton(text="Content Plan", where_add=frameMain, relx=0.55, rely=0.015, relwidth=0.15, relheight=0.05,
                       command=lambda: fileTypeSelected("Content Plan"))
        self.addButton(text="User Role Setup", where_add=frameMain, relx=0.72, rely=0.015, relwidth=0.15,
                       relheight=0.05, command=lambda: fileTypeSelected("User Role Setup"))
        self.addButton(text="Select", where_add=frameMain, relx=0.36, rely=0.33, relwidth=0.15, relheight=0.05,
                       command=lambda: selectedRoles(self.listboxList, listbox_names[0]))
        self.addButton(text="Select", where_add=frameMain, relx=0.92, rely=0.525, relwidth=0.07, relheight=0.025,
                       command=lambda: selectedMUDID(self.entryList, entry_names[0]))
        self.addButton(text="Select", where_add=frameMain, relx=0.92, rely=0.475, relwidth=0.07, relheight=0.025,
                       command=lambda: selectedUserName(self.entryList, entry_names[1]))
        self.addButton(text="Create User Role Setup", where_add=frameMain, relx=0.75, rely=0.70, relwidth=0.2, relheight=0.05,
                       command=lambda: generateFile())
        self.addButton(text="Select", where_add=frameMain, relx=0.92, rely=0.575, relwidth=0.07, relheight=0.025,
                       command=lambda: selectedAgency(self.entryList, entry_names[2]))
        self.addButton(text="Select", where_add=frameMain, relx=0.36, rely=0.6, relwidth=0.15, relheight=0.05,
                       command=lambda: selectedSecurityProfiles(self.listboxList, listbox_names[4]))
        self.addListBox(where_add=frameMain, select_mode="multiple", width=40, relx=0.1, rely=0.125,
                        contents=self.roles, name="RoleList")
        self.addListBox(where_add=frameMain, select_mode="multiple", width=40, relx=0.1, rely=0.4,
                        contents=self.securityProfiles, name="SecurityProfilesList")
        self.addEntry(where_add=frameMain, relx=0.7, rely=0.525, name="MUDID")
        self.addEntry(where_add=frameMain, relx=0.7, rely=0.475, name="userName")
        self.addEntry(where_add=frameMain, relx=0.7, rely=0.575, name="agencyName")

        # SUMMARY
        frame_summary = tk.Frame(self, bg='#c7b6a1')
        frame_summary.place(relx=0.00, rely=0.80, relwidth=1, relheight=0.15)

        self.addLabel(text="SUMMARY",  where_add=frame_summary, relx=0.05, rely=0.10, relwidth=0.15, relheight=0.1)     #6
        self.addLabel(text="File Type: ",  where_add=frame_summary, relx=0.05, rely=0.2, relwidth=0.20, relheight=0.1)  #7
        self.addLabel(text="",  where_add=frame_summary, relx=0.25, rely=0.2, relwidth=0.7, relheight=0.10)             #8
        self.addLabel(text="Roles: ",  where_add=frame_summary, relx=0.05, rely=0.30, relwidth=0.20, relheight=0.1)     #9
        self.addLabel(text="",  where_add=frame_summary, relx=0.25, rely=0.30, relwidth=0.7, relheight=0.1)             #10
        self.addLabel(text="User ID: ",  where_add=frame_summary, relx=0.05, rely=0.4, relwidth=0.20, relheight=0.1)    #11
        self.addLabel(text="",  where_add=frame_summary, relx=0.25, rely=0.4, relwidth=0.7, relheight=0.1)              #12
        self.addLabel(text="User Name: ",  where_add=frame_summary, relx=0.05, rely=0.5, relwidth=0.20, relheight=0.1)
        self.addLabel(text="",  where_add=frame_summary, relx=0.25, rely=0.5, relwidth=0.7, relheight=0.1)                      #14
        self.addLabel(text="Security profile: ", where_add=frame_summary, relx=0.05, rely=0.6, relwidth=0.20, relheight=0.1)
        self.addLabel(text="", where_add=frame_summary, relx=0.25, rely=0.6, relwidth=0.7, relheight=0.1)  # 16
        self.addLabel(text="Agency: ",  where_add=frame_summary, relx=0.05, rely=0.7, relwidth=0.20, relheight=0.1)
        self.addLabel(text="", where_add=frame_summary, relx=0.25, rely=0.7, relwidth=0.7, relheight=0.1)                       #118

        def fileTypeSelected(filetype):
            self.fileTypeSelected = filetype
            self.labelDict.get(8)['text'] = filetype
            print(filetype)


        def selectedRoles(listBoxDictList, name):
            for element in listBoxDictList:
                if element.get("Name") == name:
                    list = element.get("ListBox object")
                    self.selectedRoles = [list.get(idx) for idx in list.curselection()]
                    print(self.selectedRoles)
            self.labelDict.get(10)['text'] = self.selectedRoles

        def selectedMUDID(entryDictList, name):
            for element in entryDictList:
                if element.get("Name") == name:
                    entry = element.get("Entry object")
                    self.MUDID = entry.get()
                    print(self.MUDID)
            self.labelDict.get(12)['text'] = self.MUDID

        def selectedUserName(entryDictList, name):
            for element in entryDictList:
                if element.get("Name") == name:
                    entry = element.get("Entry object")
                    self.UserName = entry.get()
                    print(self.UserName)
            self.labelDict.get(14)['text'] = self.UserName

        def selectedSecurityProfiles(listBoxDictList, name):
            for element in listBoxDictList:
                if element.get("Name") == name:
                    list = element.get("ListBox object")
                    self.selectedSecurityProfiles = [list.get(idx) for idx in list.curselection()]
                    print(self.selectedSecurityProfiles)
            self.labelDict.get(16)['text'] = self.selectedSecurityProfiles

        def selectedAgency(entryDictList, name):
            for element in entryDictList:
                if element.get("Name") == name:
                    entry = element.get("Entry object")
                    self.agencyName = entry.get()
                    print(self.agencyName)
            self.labelDict.get(18)['text'] = self.agencyName

        def generateFile():
            user_input = {"fileType": self.fileTypeSelected,
                          "roles": self.selectedRoles,
                          "MUDID": self.MUDID,
                          "userName": self.UserName,
                          "securityProfile": self.selectedSecurityProfiles,
                          'agency': self.agencyName,
                          'markets': ['United States']}

            if user_input.get("fileType") == "Content Plan":
                contentPlan = ContentPlan(user_input).to_data_frame()
                pathfile = '_1_{}_US_content_plan.csv'.format(user_input.get("userName"))
                contentPlan.to_csv(pathfile, index=False, header=True)

            if user_input.get("fileType") == "User Role Setup":
                userRoleSetup = UsUsersUserRoleSetup(user_input).to_data_frame()
                pathfile = '_2_{}_US_user_role_setup.csv'.format(user_input.get("userName"))
                userRoleSetup.to_csv(pathfile, index=False, header=True)


    # FUNCTIONS ADDING WIDGETS
    def addLabel(self, text, where_add, relx, rely, relwidth=None, relheight=None):
        label = tk.Label(where_add, text=text)
        if relwidth and relheight is not None:
            label.place(relx=relx, rely=rely, relwidth=relwidth, relheight=relheight)
        else:
            label.place(relx=relx, rely=rely)
        dict_to_add = {len(self.labelDict): label}
        self.labelDict.update(dict_to_add)

    def addButton(self, text, where_add, relx, rely, relwidth, relheight, command=None):
        button = tk.Button(where_add, text=text)
        button.place(relx=relx, rely=rely, relwidth=relwidth, relheight=relheight)
        if command is not None:
            button.configure(command=command)
        self.buttonList.append(button)

    def addListBox(self, where_add, select_mode, width, relx, rely, contents, name):
        ListBox = tk.Listbox(where_add, selectmode=select_mode, width=width)
        ListBox.place(relx=relx, rely=rely)
        for index, each_item in enumerate(contents):
            ListBox.insert(index, each_item)
        dict_to_add = {"Name": name,
                       "ListBox object": ListBox}
        self.listboxList.append(dict_to_add)

    def addEntry(self, where_add, relx, rely, name):
        Entry = tk.Entry(where_add)
        Entry.place(relx=relx, rely=rely)
        dict_to_add = {"Name": name,
                       "Entry object": Entry}
        self.entryList.append(dict_to_add)

class Tool:

    def __init__(self, height, width):

        # MAIN PART
        self.root = tk.Tk()
        self.root.title('Internal User Creation Tool 0.4')
        self.root.iconbitmap('Pizza.ico')
        canvas = tk.Canvas(self.root, height=height, width=width)
        canvas.pack()

        # Random colour generation
        r = lambda: random.randint(0, 255)
        bg = '#%02X%02X%02X' % (r(), r(), r())

        # Placing pages
        buttonframe = tk.Frame(canvas)
        buttonframe.place(relx=0.00, rely=0.00, relwidth=1, relheight=0.05)

        frameMain = tk.Frame(canvas, bg=bg)         # #ffab45 <-- original orange colour
        frameMain.place(relx=0.00, rely=0.05, relwidth=1, relheight=1)

        p1 = InternalUsers(height = height, width = width)
        p2 = ExternalUsers(height = height, width = width)
        p3 = UsUsers(height = height, width = width)

        p1.place(in_ = frameMain, relx=0.00, rely=0.00, relwidth=1, relheight=1)
        p2.place(in_ = frameMain, relx=0.00, rely=0.00, relwidth=1, relheight=1)
        p3.place(in_ = frameMain, relx=0.00, rely=0.00, relwidth=1, relheight=1)

        b1 = tk.Button(buttonframe, text="Internal Users", command=p1.show)
        b2 = tk.Button(buttonframe, text="External Users", command=p2.show)
        b3 = tk.Button(buttonframe, text="US Users", command=p3.show)

        b1.place(in_ = buttonframe, relx=0.00, rely=0.00, relwidth=0.33333, relheight=1)
        b2.place(in_ = buttonframe, relx=0.33333, rely=0.00, relwidth=0.33333, relheight=1)
        b3.place(in_ = buttonframe, relx=0.66666, rely=0.00, relwidth=0.33333, relheight=1)


if __name__ == '__main__':
    tool = Tool(height=800, width=600)
    tool.root.mainloop()