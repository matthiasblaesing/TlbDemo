
package eu.doppel_helix.jna.tlb.demo;

import com.sun.jna.Pointer;
import com.sun.jna.platform.win32.COM.COMInvoker;
import com.sun.jna.platform.win32.COM.Dispatch;
import com.sun.jna.platform.win32.COM.util.Factory;
import com.sun.jna.platform.win32.Ole32;
import com.sun.jna.platform.win32.WTypes;
import com.sun.jna.platform.win32.WTypes.BSTR;
import eu.doppel_helix.jna.tlb.shdocvw1.IWebBrowser2;
import eu.doppel_helix.jna.tlb.shdocvw1.InternetExplorer;

/**
 * Internet Explorer Demo 1.
 *
 * <p>Open an internet explorer instance, and navigate to a sample page. Display
 * that for 5 seconds and then  force a reload.</p>
 */
public class InternetExplorerCallDemo {

    private static final int REFRESH_NORMAL = 0;
    private static final int REFRESH_IFEXPIRED = 1;
    private static final int REFRESH_COMPLETELY = 3;

    public static void main(String[] args) throws Exception {
        // Initialize COM Subsystem
        Ole32.INSTANCE.CoInitializeEx(Pointer.NULL, Ole32.COINIT_MULTITHREADED);
        // Initialize Factory for COM object creation
        Factory fact = new Factory();

        try {
            InternetExplorer ie = fact.createObject(InternetExplorer.class);
            IWebBrowser2 iw2 = ie.queryInterface(IWebBrowser2.class);
            
            ie.setVisible(Boolean.TRUE);

//            ie.Navigate("http://www.doppel-helix.eu/test.php", VARIANT_MISSING, VARIANT_MISSING, VARIANT_MISSING, VARIANT_MISSING);

            class COMInvokerMod extends COMInvoker {

                @Override
                public void _invokeNativeVoid(int vtableId, Object[] args) {
                    super._invokeNativeVoid(vtableId, args);
                }

                @Override
                public Object _invokeNativeObject(int vtableId, Object[] args, Class<?> returnType) {
                    return super._invokeNativeObject(vtableId, args, returnType);
                }

                @Override
                public int _invokeNativeInt(int vtableId, Object[] args) {
                    return super._invokeNativeInt(vtableId, args);
                }
                
            }

            Pointer basePointer = ((Dispatch)iw2.getRawDispatch()).getPointer();
            COMInvokerMod mod = new COMInvokerMod();
            mod.setPointer(basePointer);
            
            BSTR url = new WTypes.BSTR("http://www.heise.de");
            mod._invokeNativeVoid(11, new Object[]{basePointer, url.getPointer(), null, null, null, null});
            System.out.println(url.getValue());

            Thread.sleep(5 * 1000);

//            ie.Refresh2(REFRESH_COMPLETELY);

//            Thread.sleep(5 * 1000);

            ie.Quit();
        } finally {
            fact.disposeAll();
            Ole32.INSTANCE.CoUninitialize();
        }
    }

}
