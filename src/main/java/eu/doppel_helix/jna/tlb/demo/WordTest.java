package eu.doppel_helix.jna.tlb.demo;

import com.sun.jna.platform.win32.COM.util.Factory;
import static com.sun.jna.platform.win32.Variant.VARIANT.VARIANT_MISSING;
import eu.doppel_helix.jna.tlb.word8.Application;
import eu.doppel_helix.jna.tlb.word8.Document;
import eu.doppel_helix.jna.tlb.word8.PageSetup;

public class WordTest {

    public static void main(String[] args) throws InterruptedException {
        // Initialize COM Subsystem
//        Ole32.INSTANCE.CoInitializeEx(Pointer.NULL, Ole32.COINIT_MULTITHREADED);
        // Initialize Factory for COM object creation
//        ObjectFactory fact = new ObjectFactory();

        Factory fact = new Factory();

        try {

            Application appX = fact.createObject(Application.class);

//            appX.setVisible(Boolean.TRUE);
            
            Document doc = appX.getDocuments().Add(VARIANT_MISSING, VARIANT_MISSING, VARIANT_MISSING, VARIANT_MISSING);
            PageSetup ps = doc.getPageSetup();
            
            System.out.println(ps.getFirstPageTray());
            System.out.println(ps.getProperty(Long.class, "FirstPageTray"));
            
            ps.setProperty("FirstPageTray", 0);
            
            System.out.println(ps.getProperty(Long.class, "FirstPageTray"));
            System.out.println(ps.getFirstPageTray());
            
            appX.invokeMethod(Void.class, "Quit");

        } finally {
            fact.disposeAll();
            fact.getComThread().terminate(10 * 1000);
//            Ole32.INSTANCE.CoUninitialize();
        }
    }
}
