package eu.doppel_helix.jna.tlb.demo;

import com.sun.jna.Pointer;
import com.sun.jna.platform.win32.COM.util.ObjectFactory;
import com.sun.jna.platform.win32.OaIdl.SAFEARRAY;
import com.sun.jna.platform.win32.Ole32;
import eu.doppel_helix.jna.tlb.wia1.CommonDialog;
import eu.doppel_helix.jna.tlb.wia1.IImageFile;
import eu.doppel_helix.jna.tlb.wia1.WiaDeviceType;
import eu.doppel_helix.jna.tlb.wia1.WiaImageBias;
import eu.doppel_helix.jna.tlb.wia1.WiaImageIntent;
import java.io.File;

public class WIATest {

    public static void main(String[] args) {
        Ole32.INSTANCE.CoInitializeEx(Pointer.NULL, Ole32.COINIT_MULTITHREADED);
        ObjectFactory of = new ObjectFactory();
        try {
//            DeviceManager dm = of.createObject(DeviceManager.class);
//            IDeviceInfos infos = dm.getDeviceInfos();
//            for(int i = infos.getCount(); i > 0; i--) {
//                IDeviceInfo idi = infos.getItem(i);
//                IProperties props = idi.getProperties();
//                System.out.println(idi.getDeviceID());
//                System.out.println(idi.getType());
//                for(int j = props.getCount(); j > 0; j--) {
//                    IProperty prop = props.getItem(j);
//                    System.out.println(prop.getName() + " => " + prop.getValue());
//                }
//            }

            CommonDialog cd = of.createObject(CommonDialog.class);
            IImageFile iif = cd.ShowAcquireImage(
                    WiaDeviceType.ScannerDeviceType,
                    WiaImageIntent.UnspecifiedIntent,
                    WiaImageBias.MaximizeQuality,
                    "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}",
                    Boolean.TRUE,
                    Boolean.TRUE,
                    Boolean.FALSE);

            new File("c:/temp/test.png").delete();
            iif.SaveFile("c:/temp/test.png");
            SAFEARRAY sa = (SAFEARRAY) iif.getFileData().getBinaryData();
            System.out.println(sa.getVarType());
            Pointer p = sa.accessData();
            System.out.println(new String(p.getByteArray(0, 4)));
            sa.unaccessData();
//            int count = vect.getCount();
//            for (int i = 1; i <= count; i++) {
//                long val = ((Number) vect.getItem(i)).longValue();
//                int alpha = (int) ((val >>> 24) & 0xFF);
//                int red = (int) ((val >>> 16) & 0xFF);
//                int green = (int) ((val >>> 8) & 0xFF);
//                int blue = (int) ((val >>> 0) & 0xFF);
//                System.out.println(String.format("(%d, %d, %d)", red, green, blue));
//                break;
//            }
        } finally {
            of.disposeAll();
            Ole32.INSTANCE.CoUninitialize();
        }
    }
}
