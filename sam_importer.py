import sys
import os
from inspect import getsourcefile
from os.path import abspath

# prevent python from writing *.pyc files / __pycache__ folders
sys.dont_write_bytecode = True

path_app = abspath(getsourcefile(lambda: 0))[:-15]
if path_app not in sys.path:
    sys.path.append(path_app)

import tkinter as tk
import xml.etree.ElementTree as etree
from xlwt import Workbook
from xlrd import open_workbook
from preconfigured_ttk_widgets import *
# import httplib2

class SAMImporter(MainWindow):
    
    def __init__(self, path_app):
        super().__init__()
        self.main_path = path_app
        self.request_path = path_app + 'request'
        self.response_path = path_app + 'response'
        
        self.obj_classes = {
        'nodes': ('netw.NetworkElement',),
        'links': ('netw.PhysicalLink', 'netw.OpticalLink'),
        'ports': ('equipment.PhysicalPort', 'equipment.LogicalPort'),
        'interfaces': ('netw.LogicalInterface',),
        'power': ('optical.PowerSpecifics',)
        }
        
        self.obj_properties = {
        
        'nodes': (
        'siteId',
        'objectFullName',
        'name',
        'mgmtIpAddrType',
        'ipAddress',
        'baseMacAddress',
        'outOfBandAddress',
        'inBandSystemAddress',
        'inBandL3ManagementIf',
        'sysDescription',
        'systemAddress',
        'location',
        'coordinates',
        'chassisType',
        'productType',
        'sysObjectId',
        'latitudeInDegrees',
        'longitudeInDegrees',
        'neState',
        'locationId',
        'olcState',
        'deploymentState'
        ),
        
        'links': (
        'objectFullName',
        'name',
        'deploymentState',
        'description',
        'displayedName',
        'endpointAPointer',
        'endpointBPointer',
        'endPointASiteId',
        'endPointASiteName',
        'endPointASiteType',
        'endPointAChassisId',
        'endPointAChassisType',
        'endPointAPortId',
        'endPointBSiteType',
        'endPointBSiteId',
        'endPointBSiteName',
        'endPointBChassisId',
        'endPointBChassisType',
        'endPointBPortId',
        'endpointBSubnetPointer',
        'linkDiscoveredFrom',
        'linkType',
        'physicalLinkScope',
        'physicalinkPointers',
        'underlyingEndpointAPhysLinkPointer',
        'underlyingEndpointBPhysLinkPointer',
        'isLagMember',
        'endPointALagMembershipId',
        'endPointBLagMembershipId'
        ),
        
        'ports': (
        'siteName',
        'siteId',
        'portId',
        'portName',
        'displayedName',
        'description',
        'objectFullName',
        'name',
        'deploymentState',
        'isEquipped',
        'isEquipmentInserted',
        'isLinkUp',
        'administrativeState',
        'equipmentState',
        'olcState',
        'shelfId',
        'speed',
        'actualSpeed',
        'encapType',
        'macAddress',
        'operationalMTU',
        'portCategory',
        'portClass',
        'cardSlotId',
        'daughterCardSlotId',
        'isEnabled',
        'isPrimaryLagMember',
        'lagMembershipId',
        'portAccessDescription',
        'vlan',
        'vplsMode'
        ),
        
        'interfaces': (
        'application',
        'displayedName',
        'objectFullName',
        'domain',
        'nodeId',
        'nodeName',
        'deploymentState',
        'portId',
        'portName',
        'provisionedMtu',
        'actualMtu',
        'routerId',
        'routerName',
        'terminatedObjectId',
        'terminatedObjectName',
        'terminatedObjectPointer',
        'terminatedPortClassName', 
        'terminatedPortPointer'
        ),
        
        'power': (
        'siteName',
        'siteId',
        'name',
        'displayedName',
        'objectFullName',
        'portId',
        'cardSlotId',
        'portRole',
        'portType',
        # connected to an internal or external port
        'apFarEndType',
        'rmnPortSignalPowerOut',
        'rmnPortTotalPowerIn',
        'wkPortChannelEgressPower',
        'wkPortChannelIngressPower',
        'wkPortNwPowerIn',
        'wkPortNwPowerOut'
        )}
                    
        # title of the main window
        self.title('5620 SAM Importer')
        
        # SAM parameters, initalized with the 'default_parameters' file
        with open(self.main_path + 'default_parameters.txt', 'r') as file:
            for line in file:
                parameter, value = line.replace(' ', '').split(':')
                setattr(self, parameter, value)
        
        # main menu
        menubar = Menu(self)
        upper_menu = Menu(menubar)

        # first menu entry: SAM parameters
        parameters = MenuEntry(upper_menu)
        parameters.text = 'Parameters'
        parameter_window = Parameters(self)
        parameters.command = lambda: parameter_window.deiconify()
        
        upper_menu.create_menu()
        menubar.add_cascade(label='Options', menu=upper_menu)
        self.config(menu=menubar)
        
        # main window
        main_frame = MainFrame(self)
        main_frame.pack()
        
class MainFrame(CustomFrame):
    
    def __init__(self, master):
        super().__init__()
        self.ms = master
        
        # label frame
        lf = Labelframe(self)
        lf.text = 'SAM Importer'
                                                        
        # object listbox and associated scrollbar
        objects_listbox = Listbox(self, width=15, height=7)   
        yscroll = Scrollbar(self)
        yscroll.command = objects_listbox.yview
        objects_listbox.configure(yscrollcommand=yscroll.set)
        
        for obj in master.obj_classes:
            objects_listbox.insert(obj)
            
        # display the SAM response
        sam_response = Text(self, width=50, height=10)
        
        # button to send requests
        button_send = Button(self, width=13)
        button_send.text = 'Send requests'
        button_send.command = self.HTTP_post_request
        
        # button to convert xml files to xls
        button_convert = Button(self, width=13)
        button_convert.text = 'XLS conversion'
        button_convert.command = self.XLS_conversion
        
        # labelframe grid
        lf.grid(0, 0)
        button_send.grid(4, 0, in_=lf)
        button_convert.grid(5, 0, in_=lf)
        objects_listbox.grid(0, 0, 3, in_=lf)
        sam_response.grid(0, 2, 6, 3, padx=10, in_=lf)
        yscroll.grid(0, 1, 3, in_=lf)

    def HTTP_post_request(self):
        httplib2.debuglevel = 1
        http = httplib2.Http()
        
        for file in os.listdir(self.request_path):
            path_file = self.request_path + '\\' + file
            request = open(path_file, 'rb').read()

        try: 
            address = "http://%s:%s/xmlapi/invoke" % (self.SAM_IP, self.SAM_port)
            response, content = http.request(address, 'POST', body=request)
            
            # write the output xml file with the SAM content
            output_file = open(path_file, 'w')
            output_file.write(str(content)[2:-1])
            output_file.close()
        
            # display the SAM HTTP request response
            self.sam_response.text = response
            
        except (
                ConnectionRefusedError, 
                OSError, 
                httplib2.ServerNotFoundError
                ) as error:
            self.sam_response.text = str(error)
            
    def XLS_conversion(self):
        for file in os.listdir(self.ms.response_path):
            path_file = self.ms.response_path + '\\' + file
            obj_type, file_type = file.split('.')
            path_output = path_file + obj_type + '.xls'
            
            if file_type == 'xml':
                tree = etree.parse(path_file)
                book = Workbook()
                
                xls_sheet = book.add_sheet('SAM data', cell_overwrite_ok=True)
                for col_id in range(50):
                    # 25 characters wide
                    xls_sheet.col(col_id).width = 256 * 25
                
                for idx, property in enumerate(self.ms.obj_properties[obj_type]):
                    xls_sheet.write(0, idx, property)                    
                
                k = 0
                for node in tree.iter():
                    tag = node.tag[12:]
                    if tag in self.ms.obj_classes[obj_type]:
                        k += 1
                    else:
                        col_id = self.ms.obj_properties[obj_type].index(tag)
                        xls_sheet.write(k, q, node.text)
                        
                book.save(path_output)
        
class Parameters(CustomTopLevel):

    def __init__(self, master):
        super().__init__()
        self.ms = master
        
        # label frame
        parameters_lf = Labelframe(self)
        parameters_lf.text = 'Parameters'
        
        # label + associated entry
        sam_ip_label = Label(self)
        sam_ip_label.text = 'SAM IP :'
        self.sam_ip_entry = Entry(self, width=13)
        self.sam_ip_entry.text = self.ms.SAM_IP
        
        self.sam_port = tk.IntVar()
        self.sam_port.set(int(self.ms.SAM_port)) 
        standard_port = Radiobutton(self, variable=self.sam_port, value=8080)
        standard_port.text = 'Port 8080'
        https_port = Radiobutton(self, variable=self.sam_port, value=8443)
        https_port.text = 'Port 8443'
        
        parameters_lf.grid(0, 0)
        sam_ip_label.grid(0, 0, in_=parameters_lf)
        self.sam_ip_entry.grid(0, 1, in_=parameters_lf)
        standard_port.grid(1, 0, in_=parameters_lf)
        https_port.grid(1, 1, in_=parameters_lf)
        
        # save options when closing the parameters window
        self.protocol('WM_DELETE_WINDOW', self.save_parameters)
        self.withdraw()
        
    def save_parameters(self):
        with open(self.ms.main_path + 'default_parameters.txt', 'w') as file:
            file.write('SAM_IP: ' + self.sam_ip_entry.get())
            file.write('SAM_port: ' + str(self.sam_port.get()))
        self.withdraw()

if __name__ == '__main__':
    SAM_importer = SAMImporter(path_app)
    SAM_importer.mainloop()