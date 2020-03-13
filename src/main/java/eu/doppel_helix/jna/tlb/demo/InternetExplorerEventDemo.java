
package eu.doppel_helix.jna.tlb.demo;

import com.sun.jna.Pointer;
import com.sun.jna.platform.win32.COM.util.IComEventCallbackCookie;
import com.sun.jna.platform.win32.COM.util.IDispatch;
import com.sun.jna.platform.win32.COM.util.ObjectFactory;
import com.sun.jna.platform.win32.Ole32;
import com.sun.jna.platform.win32.Variant.VARIANT;
import eu.doppel_helix.jna.tlb.shdocvw1.DWebBrowserEvents2Listener;
import eu.doppel_helix.jna.tlb.shdocvw1.DWebBrowserEvents2ListenerHandler;
import eu.doppel_helix.jna.tlb.shdocvw1.InternetExplorer;

/**
 * Internet Explorer Demo 2
 *
 * <p>
 * Demonstrace recoding of events. An internet explorer instance is opened and
 * Navigate Events are printed. When the quit event is intercepted, the monitor
 * is shutdown.</p>
 */
public class InternetExplorerEventDemo {

    public static void main(String[] args) throws Exception {
        // Initialize COM Subsystem
        Ole32.INSTANCE.CoInitializeEx(Pointer.NULL, Ole32.COINIT_MULTITHREADED);
        // Initialize Factory for COM object creation
        ObjectFactory fact = new ObjectFactory();

        try {
            InternetExplorer ie = fact.createObject(InternetExplorer.class);
            ie.setVisible(Boolean.TRUE);

            class DWebBrowserEvents2_Listener extends DWebBrowserEvents2ListenerHandler {

                @Override
                public void errorReceivingCallbackEvent(String message, Exception exception) {
                    System.err.println(message);
                    exception.printStackTrace(System.err);
                }

                @Override
                public void BeforeNavigate2(IDispatch pDisp, Object URL, Object Flags, Object TargetFrameName, Object PostData, Object Headers, VARIANT Cancel) {
                    if(URL instanceof String) {
                        System.out.println("Before navigate: " + ((String) URL));
                    }
                }
                
                @Override
                public void NavigateComplete2(IDispatch pDisp, Object URL) {
                    if(URL instanceof String) {
                        System.out.println("Navigate done: " + ((String) URL));
                    }
                }

                volatile Boolean quitCalled = false;

                @Override
                public void OnQuit() {
                    quitCalled = true;
                }
            }
            
            class DWebBrowserEvents2_Listener2 extends DWebBrowserEvents2ListenerHandler {

                @Override
                public void errorReceivingCallbackEvent(String message, Exception exception) {
                    System.err.println(message);
                    exception.printStackTrace(System.err);
                }

                @Override
                public void BeforeNavigate2(IDispatch pDisp, Object URL, Object Flags, Object TargetFrameName, Object PostData, Object Headers, VARIANT Cancel) {
                    if(URL instanceof String) {
                        System.out.println("Before navigate 2: " + ((String) URL));
                    }
                }
                
                @Override
                public void NavigateComplete2(IDispatch pDisp, Object URL) {
                    if(URL instanceof String) {
                        System.out.println("Navigate done 2: " + ((String) URL));
                    }
                }

                volatile Boolean quitCalled = false;

                @Override
                public void OnQuit() {
                    quitCalled = true;
                }
            }

            DWebBrowserEvents2_Listener listener = new DWebBrowserEvents2_Listener();
            DWebBrowserEvents2_Listener2 listener2 = new DWebBrowserEvents2_Listener2();
            
            ie.advise(DWebBrowserEvents2Listener.class, listener);
            IComEventCallbackCookie cookie = ie.advise(DWebBrowserEvents2Listener.class, listener2);
            
            int count = 0;
            while(! listener.quitCalled) {
                count++;
                Thread.sleep(500);
                if(count == 10) {
                    ie.unadvise(DWebBrowserEvents2Listener.class, cookie);
                }
            }
            
            System.out.println("QUIT");
        } finally {
            fact.disposeAll();
            Ole32.INSTANCE.CoUninitialize();
        }
    }

}
