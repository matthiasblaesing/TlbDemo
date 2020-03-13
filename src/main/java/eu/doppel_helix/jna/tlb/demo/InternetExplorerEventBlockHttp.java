
package eu.doppel_helix.jna.tlb.demo;

import com.sun.jna.Pointer;
import com.sun.jna.platform.win32.COM.util.AbstractComEventCallbackListener;
import com.sun.jna.platform.win32.COM.util.Factory;
import com.sun.jna.platform.win32.COM.util.IDispatch;
import com.sun.jna.platform.win32.OaIdl;
import com.sun.jna.platform.win32.OaIdl.VARIANT_BOOL;
import com.sun.jna.platform.win32.Ole32;
import com.sun.jna.platform.win32.Variant.VARIANT;
import eu.doppel_helix.jna.tlb.shdocvw1.DWebBrowserEvents2;
import eu.doppel_helix.jna.tlb.shdocvw1.DWebBrowserEvents2Listener;
import eu.doppel_helix.jna.tlb.shdocvw1.InternetExplorer;

/**
 * Internet Explorer Demo 3
 *
 * <p>Demonstrate intercept loading of URLs that are loaded over http.</p>
 */
public class InternetExplorerEventBlockHttp {

    public static void main(String[] args) throws Exception {
        // Initialize COM Subsystem
        Ole32.INSTANCE.CoInitializeEx(Pointer.NULL, Ole32.COINIT_MULTITHREADED);
        // Initialize Factory for COM object creation
        Factory fact = new Factory();

        try {
            InternetExplorer ie = fact.createObject(InternetExplorer.class);
            ie.setVisible(Boolean.TRUE);

            class DWebBrowserEvents2_Listener extends AbstractComEventCallbackListener implements DWebBrowserEvents2Listener {

                @Override
                public void errorReceivingCallbackEvent(String message, Exception exception) {
                    System.err.println(message);
                    exception.printStackTrace(System.err);
                }

                @Override
                public void BeforeNavigate2(IDispatch pDisp, Object URL, Object Flags, Object TargetFrameName, Object PostData, Object Headers, VARIANT Cancel) {
                    if(URL instanceof String) {
                        String url = ((String) URL);
                        boolean block = url.startsWith("http") && (! url.startsWith("https"));
                        System.out.println("Before navigate: " + url + (block ? " BLOCKED " : ""));
                        ((OaIdl.VARIANT_BOOLByReference) Cancel.getValue()).setValue(new VARIANT_BOOL(block));
                    }
                }

                volatile Boolean quitCalled = false;

                @Override
                public void OnQuit() {
                    quitCalled = true;
                }

                @Override
                public void StatusTextChange(String Text) {

                }

                @Override
                public void ProgressChange(Integer Progress, Integer ProgressMax) {

                }

                @Override
                public void CommandStateChange(Integer Command, Boolean Enable) {

                }

                @Override
                public void DownloadBegin() {

                }

                @Override
                public void DownloadComplete() {

                }

                @Override
                public void TitleChange(String Text) {

                }

                @Override
                public void PropertyChange(String szProperty) {

                }

                @Override
                public void NewWindow2(VARIANT ppDisp, VARIANT Cancel) {

                }

                @Override
                public void NavigateComplete2(IDispatch pDisp, Object URL) {

                }

                @Override
                public void DocumentComplete(IDispatch pDisp, Object URL) {

                }

                @Override
                public void OnVisible(Boolean Visible) {

                }

                @Override
                public void OnToolBar(Boolean ToolBar) {

                }

                @Override
                public void OnMenuBar(Boolean MenuBar) {

                }

                @Override
                public void OnStatusBar(Boolean StatusBar) {

                }

                @Override
                public void OnFullScreen(Boolean FullScreen) {

                }

                @Override
                public void OnTheaterMode(Boolean TheaterMode) {

                }

                @Override
                public void WindowSetResizable(Boolean Resizable) {

                }

                @Override
                public void WindowSetLeft(Integer Left) {

                }

                @Override
                public void WindowSetTop(Integer Top) {

                }

                @Override
                public void WindowSetWidth(Integer Width) {

                }

                @Override
                public void WindowSetHeight(Integer Height) {

                }

                @Override
                public void WindowClosing(Boolean IsChildWindow, VARIANT Cancel) {

                }

                @Override
                public void ClientToHostWindow(VARIANT CX, VARIANT CY) {

                }

                @Override
                public void SetSecureLockIcon(Integer SecureLockIcon) {

                }

                @Override
                public void FileDownload(Boolean ActiveDocument, VARIANT Cancel) {

                }

                @Override
                public void NavigateError(IDispatch pDisp, Object URL, Object Frame, Object StatusCode, VARIANT Cancel) {

                }

                @Override
                public void PrintTemplateInstantiation(IDispatch pDisp) {

                }

                @Override
                public void PrintTemplateTeardown(IDispatch pDisp) {

                }

                @Override
                public void UpdatePageStatus(IDispatch pDisp, Object nPage, Object fDone) {

                }

                @Override
                public void PrivacyImpactedStateChange(Boolean bImpacted) {

                }

                @Override
                public void NewWindow3(VARIANT ppDisp, VARIANT Cancel, Integer dwFlags, String bstrUrlContext, String bstrUrl) {

                }

                @Override
                public void SetPhishingFilterStatus(Integer PhishingFilterStatus) {

                }

                @Override
                public void WindowStateChanged(Integer dwWindowStateFlags, Integer dwValidFlagsMask) {

                }

                @Override
                public void NewProcess(Integer lCauseFlag, IDispatch pWB2, VARIANT Cancel) {

                }

                @Override
                public void ThirdPartyUrlBlocked(Object URL, Integer dwCount) {
                }

                @Override
                public void RedirectXDomainBlocked(IDispatch pDisp, Object StartURL, Object RedirectURL, Object Frame, Object StatusCode) {
                }

                @Override
                public void BeforeScriptExecute(IDispatch pDispWindow) {
                }

                @Override
                public void WebWorkerStarted(Integer dwUniqueID, String bstrWorkerLabel) {
                }

                @Override
                public void WebWorkerFinsihed(Integer dwUniqueID) {
                }
                
                
            }

            DWebBrowserEvents2_Listener listener = new DWebBrowserEvents2_Listener();
            
            ie.advise(DWebBrowserEvents2Listener.class, listener);
            
            while(! listener.quitCalled) {
                Thread.sleep(500);
            }
            
            System.out.println("QUIT");
        } finally {
            fact.disposeAll();
            Ole32.INSTANCE.CoUninitialize();
        }
    }

}
