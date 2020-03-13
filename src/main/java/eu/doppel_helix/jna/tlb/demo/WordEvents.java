package eu.doppel_helix.jna.tlb.demo;

import com.sun.jna.Pointer;
import com.sun.jna.platform.win32.COM.util.Factory;
import com.sun.jna.platform.win32.COM.util.IComEventCallbackCookie;
import com.sun.jna.platform.win32.Ole32;
import static com.sun.jna.platform.win32.Variant.VARIANT.VARIANT_MISSING;
import eu.doppel_helix.jna.tlb.word8.Application;
import eu.doppel_helix.jna.tlb.word8.ApplicationEvents4Listener;
import eu.doppel_helix.jna.tlb.word8.ApplicationEvents4ListenerHandler;
import eu.doppel_helix.jna.tlb.word8.Document;

public class WordEvents {

    public static void main(String[] args) throws InterruptedException {
        // Initialize COM Subsystem
        Ole32.INSTANCE.CoInitializeEx(Pointer.NULL, Ole32.COINIT_MULTITHREADED);
        // Initialize Factory for COM object creation
        Factory fact = new Factory();

        try {
            class ApplicatonEventsHandler extends ApplicationEvents4ListenerHandler {

                public volatile boolean quit = false;
                
                @Override
                public void errorReceivingCallbackEvent(String string, Exception excptn) {
                    System.out.println("Error: " + string);
                }

                @Override
                public void Startup() {
                    System.out.println("Startup");
                }

                @Override
                public void Quit() {
                    quit = true;
                }

                @Override
                public void DocumentChange() {
                    System.out.println("DocumentChange");
                }
                
            }
            
            ApplicatonEventsHandler handler = new ApplicatonEventsHandler();
            
            Application appX = fact.createObject(Application.class);
            
            Thread.sleep(2 * 1000);
            
            IComEventCallbackCookie cookie = appX.advise(ApplicationEvents4Listener.class, handler);
            
            Document doc = appX.getDocuments().Add(VARIANT_MISSING, VARIANT_MISSING, VARIANT_MISSING, VARIANT_MISSING);
            doc.getParagraphs().Item(1).getRange().setText("Test text");
            doc.Close(Boolean.FALSE, VARIANT_MISSING, VARIANT_MISSING);
            
            appX.unadvise(ApplicationEvents4Listener.class, cookie);
            
            appX.Quit(Boolean.FALSE, VARIANT_MISSING, VARIANT_MISSING);
        } finally {
            fact.disposeAll();
            Ole32.INSTANCE.CoUninitialize();
        }
    }
}

