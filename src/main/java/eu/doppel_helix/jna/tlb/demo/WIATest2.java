package eu.doppel_helix.jna.tlb.demo;

import com.sun.jna.Pointer;
import com.sun.jna.platform.win32.COM.util.ObjectFactory;
import com.sun.jna.platform.win32.Ole32;
import eu.doppel_helix.jna.tlb.wia1.DeviceManager;
import eu.doppel_helix.jna.tlb.wia1.IDevice;
import eu.doppel_helix.jna.tlb.wia1.IDeviceCommand;
import eu.doppel_helix.jna.tlb.wia1.IDeviceCommands;
import eu.doppel_helix.jna.tlb.wia1.IDeviceInfo;
import eu.doppel_helix.jna.tlb.wia1.IDeviceInfos;
import eu.doppel_helix.jna.tlb.wia1.IItem;
import eu.doppel_helix.jna.tlb.wia1.IItems;
import eu.doppel_helix.jna.tlb.wia1.IProperties;
import eu.doppel_helix.jna.tlb.wia1.IProperty;

public class WIATest2 {

    public static void main(String[] args) {
        Ole32.INSTANCE.CoInitializeEx(Pointer.NULL, Ole32.COINIT_MULTITHREADED);
        ObjectFactory of = new ObjectFactory();
        try {
            DeviceManager dm = of.createObject(DeviceManager.class);
            IDeviceInfos infos = dm.getDeviceInfos();
            for(int i = infos.getCount(); i > 0; i--) {
                System.out.println("--------------------------");
                IDeviceInfo idi = infos.getItem(i);
//                IProperties props = idi.getProperties();
//                for(int j = props.getCount(); j > 0; j--) {
//                    IProperty prop = props.getItem(j);
//                    System.out.println(prop.getName() + " => " + prop.getValue());
//                }
                
                IDevice device = idi.Connect();
                IProperties deviceProperties = device.getProperties();
                for (int j = deviceProperties.getCount(); j > 0; j--) {
                    IProperty prop = deviceProperties.getItem(j);
                    System.out.println(prop.getName() + " => " + prop.getValue());
                }
                
                
                IDeviceCommands commands = device.getCommands();
                for(int j = commands.getCount(); j > 0; j--) {
                    IDeviceCommand cmd = commands.getItem(j);
                    System.out.println("CMD: (" + cmd.getCommandID() + ") -> " + cmd.getName());
                    System.out.println("\t" + cmd.getDescription());
                }
                
                IItems items = device.getItems();
                for(int j = items.getCount(); j > 0; j--) {
                    System.out.println("=================");
                    IItem item = items.getItem(j);
                    System.out.println(item.getProperties().getItem("Item Name").getValue());
                }
            }

            
        } finally {
            of.disposeAll();
            Ole32.INSTANCE.CoUninitialize();
        }
    }
}
