import requests
import time
import openpyxl
import warnings


class Bulk:
    def __init__(self, mydict):
        self.url = mydict['url']
        self.user = mydict['API user']
        self.applicationpassword = mydict['applicationpassword']
        self.endpoint = '/tas/api/incidents'
        self.incidents = []
        self.changes = []
        self.decision_object = {} # duplicatE?
        self.possibilities = ["caller","branch", "short description", "category", "subcategory", "object", "operator", "operator group", "processing status", "supplier"]
        self.checker = []

        
    def self_validate(self):
        if len(self.applicationpassword) != 29:
            print("Application password is invalid")
            return False
        if not self.url.startswith("https://"):
            print("URL is invalid")
            return False
        try:
            mytestrequest = requests.get(self.url + self.endpoint,auth = (self.user, self.applicationpassword))
            if mytestrequest.status_code > 300:
                print("Initial check failed. Something was wrong with the request.: " , mytestrequest.status_code)
            else:
                print()
                print("Initial check succeeded.")
                print()
        except:
            print("Initial check failed. Something was wrong with the request.")
            return False
    def examine_iterable(self):

        myfilename = input("Please enter the name of the file with extension. The file has to be in the same folder as the script.: ")
        
        while myfilename.find(".xlsx") == -1:
            print()
            print("Wrong file! You should provide an excel file (.xlsx) with the incidents column in it.")
            print()
            myfilename = input("Please enter the name of the file with extension (a .xlsx file is expected). The file has to be in the same folder as the script.: ")
            
        try:           
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(filename = myfilename)
            warnings.simplefilter("default")

        except:
            print()
            print('Error while opening the file. Please check if the file exists.')
            print()
            return False

        
        try:
            warnings.simplefilter("ignore")
            first_sheet = wb.get_sheet_names()[0]
            worksheet = wb.get_sheet_by_name(first_sheet)
            print()
            whichColumn = input("Which column contains the incident numbers? ")
            print()
            for col in worksheet[whichColumn]:
                self.incidents.append(col.value)
            self.incidents.remove(self.incidents[0])
            for inc in self.incidents:
                int(inc[-3:])
                
            warnings.simplefilter("default")
            return True
    
        except:
            print()
            print('Error while reading the file. Please check if you provided the correct column identifier where the incidents are. If the incidents can be found in column "B", please enter B')
            print()
            return False
    def prepare_body(self):
        print()
        print('Possible fields to be changed: ')
        print("caller")
        print("branch")
        print("short description")
        print("category")
        print("subcategory")
        print("object")
        print("operator")
        print("operator group")
        print("processing status")
        print("supplier")
        
        toBeChanged = input("Please enter the fields you wish to change on the incidents, separated by comma's & press ENTER afterwards (example: category,subcategory,supplier ENTER): ")
       
        change_inputs = toBeChanged.split(',')

        for change_input in change_inputs:
            self.changes.append(change_input.strip())


        final_object = {}

        
        print()
        print("Fields {caller & operator} work with dynamic names only. Dynamic name = First Name, Surname ... Do not forget the comma!")
        print("Example: If my operator has the dynamic name 'Sándor,' (note that the last name of Sándor is missing), the input should be exactly the same, thus: Sándor,")
        print()
        for change in self.changes:
            if change != '':
                if change in self.possibilities:
                    
                    if change == "caller" or change == "operator":
                        self.decision_object[change] = input("Please enter what the dynamic name of the " + change + " should be: ")
                    else:
                        self.decision_object[change] = input("Please enter what the " + change + " should be: ")

        ## process each change to create a POST body object

        if 'caller' in self.decision_object:
            personrequest = requests.get(self.url + "/tas/api/persons?query=dynamicName==" +'"'+ self.decision_object['caller']+ '"' ,auth = (self.user, self.applicationpassword))
            personid = personrequest.json()[0]['id']


        if 'branch' in self.decision_object:
            branchrequest = requests.get(self.url + "/tas/api/branches?query=name==" +'"'+ self.decision_object['branch']+ '"' ,auth = (self.user, self.applicationpassword))
            branchid = branchrequest.json()[0]['id']

        if 'operator' in self.decision_object:
            operatorrequest = requests.get(self.url + "/tas/api/operators?query=dynamicName==" +'"'+ self.decision_object['operator']+ '"' ,auth = (self.user, self.applicationpassword))
            operatorid = operatorrequest.json()[0]['id']

        if 'operator group' in self.decision_object:
            operatorgrouprequest = requests.get(self.url + "/tas/api/operatorgroups?query=groupName==" +'"'+ self.decision_object['operator group']+ '"' ,auth = (self.user, self.applicationpassword))
            operatorgroupid = operatorgrouprequest.json()[0]['id']
     

        if 'supplier' in self.decision_object:
            supplierrequest = requests.get(self.url + "/tas/api/suppliers?query=name==" +'"'+ self.decision_object['supplier']+ '"' ,auth = (self.user, self.applicationpassword))
            supplierid = supplierrequest.json()[0]['id']

            
        for dec_key in self.decision_object:
            if dec_key == "caller":
                final_object["callerLookup"] = {"id": personid}
            elif dec_key == "branch":
                final_object["caller"] = {"branch": {"id": branchid}}
            elif dec_key == "short description":
                final_object["briefDescription"] = self.decision_object[dec_key]
            elif dec_key == "category":
                final_object["category"] = {"name": self.decision_object[dec_key]}
            elif dec_key == "subcategory":
                final_object["subcategory"] = {"name": self.decision_object[dec_key]}
            elif dec_key == "object":
                final_object["object"] = {"name": self.decision_object[dec_key]}
            elif dec_key == "operator":
                final_object["operator"] = {"id": operatorid}
            elif dec_key == "operator group":
                final_object["operatorGroup"] = {"id": operatorgroupid}
            elif dec_key == "supplier":
                final_object["supplier"] = {"id": supplierid}
            elif dec_key == "processing status":
                final_object["processingStatus"] = {"name": self.decision_object[dec_key]}
        return final_object
        
    def send_request(self, inc, patch_body):
        r = requests.patch(self.url + '/tas/api/incidents/number/' + inc, json = patch_body, auth = (self.user, self.applicationpassword))
        if r.status_code < 300:
            self.checker.append("done")
            print(inc, " done")
        else:
            print("Something went wrong with the update.", r.status_code, r.content)

    def give_feedback(self):
        print(round(len(self.checker) / len(self.incidents) * 100), "% ", "(", len(self.checker), "/", len(self.incidents),") of the incidents have been updated")





if __name__ == "__main__":
    
    parameters = {'API user':'', 'applicationpassword':'', 'url':''}

    for i in parameters.keys():
        parameters[i] = input("Please enter the " + i + ": ")
        while parameters[i] == "" or parameters[i] == " ":
            print("Field cannot be empty")
            parameters[i] = input("Please enter the " + i + ": ")

    myinputbulkobject = Bulk(parameters)


    if myinputbulkobject.self_validate() != False:
        
        if myinputbulkobject.examine_iterable() == True:
                post_obj = myinputbulkobject.prepare_body()
                print("Incidents are being updated...")

                for inc in myinputbulkobject.incidents:
                    mypatchreq = myinputbulkobject.send_request(inc, post_obj)
                    
                myinputbulkobject.give_feedback()
        else:
            print("Couldn't process the file you provided. Please re-run the script and make sure you follow the instructions.")
            


