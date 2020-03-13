
package eu.doppel_helix.jna.tlb.demo;

import com.google.zxing.BarcodeFormat;
import com.google.zxing.MultiFormatWriter;
import com.google.zxing.WriterException;
import com.google.zxing.client.j2se.MatrixToImageWriter;
import com.google.zxing.common.BitMatrix;
import com.sun.jna.Pointer;
import com.sun.jna.platform.win32.COM.util.Factory;
import com.sun.jna.platform.win32.Ole32;
import static com.sun.jna.platform.win32.Variant.VARIANT.VARIANT_MISSING;
import eu.doppel_helix.jna.tlb.office2.MsoShapeType;
import eu.doppel_helix.jna.tlb.office2.MsoTriState;
import eu.doppel_helix.jna.tlb.word8.Application;
import eu.doppel_helix.jna.tlb.word8.Document;
import eu.doppel_helix.jna.tlb.word8.Field;
import eu.doppel_helix.jna.tlb.word8.Fields;
import eu.doppel_helix.jna.tlb.word8.InlineShape;
import eu.doppel_helix.jna.tlb.word8.Range;
import eu.doppel_helix.jna.tlb.word8.Row;
import eu.doppel_helix.jna.tlb.word8.Rows;
import eu.doppel_helix.jna.tlb.word8.Shape;
import eu.doppel_helix.jna.tlb.word8.Shapes;
import eu.doppel_helix.jna.tlb.word8.Table;
import eu.doppel_helix.jna.tlb.word8.Tables;
import eu.doppel_helix.jna.tlb.word8.WdExportCreateBookmarks;
import eu.doppel_helix.jna.tlb.word8.WdExportFormat;
import eu.doppel_helix.jna.tlb.word8.WdExportItem;
import eu.doppel_helix.jna.tlb.word8.WdExportOptimizeFor;
import eu.doppel_helix.jna.tlb.word8.WdExportRange;
import eu.doppel_helix.jna.tlb.word8.WdOpenFormat;
import eu.doppel_helix.jna.tlb.word8.WdSaveOptions;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.StringWriter;
import java.math.BigDecimal;
import java.math.MathContext;
import java.math.RoundingMode;
import java.nio.file.Files;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import javax.imageio.ImageIO;
import javax.json.Json;
import javax.json.JsonObject;

public class Automation_ReplaceFormField {
    private static final double CM_TO_POINT = 72 / 2.54;
    
    public static void main(String[] args) throws IOException {
        // Prepare testdata
        Map<String,String> plainData = new HashMap<>();
        
        SimpleDateFormat germanDate = new SimpleDateFormat("dd.MM.yyyy");
        DecimalFormat euroFormat = new DecimalFormat("0.00 €", DecimalFormatSymbols.getInstance(Locale.GERMANY));
        
        BigDecimal betrag = new BigDecimal("38.75");
        betrag.setScale(2, RoundingMode.HALF_UP);
        BigDecimal skontobetrag = betrag.multiply(new BigDecimal("0.98"));
        skontobetrag = skontobetrag.setScale(2, RoundingMode.HALF_UP);
        int rechnungsnummer = 4711;
        Date datum = new Date(2016 - 1900, 3 - 1, 31);
        Date auftragDatum = new Date(2016 - 1900, 2 - 1, 2);
        Date skontoDatum = new Date(2016 - 1900, 4 - 1, 15);
        Date zahlungDatum = new Date(2016 - 1900, 4 - 1, 30);
        
        // AP_UNTERSCHRIFT is handled specially below
        plainData.put("FIRMA", "Musterfirma");
        plainData.put("KD_VORNAME", "Maxine");
        plainData.put("KD_NACHNAME", "Musterfrau");
        plainData.put("STRASSE", "Musterstraße 3");
        plainData.put("PLZ", "98765");
        plainData.put("ORT", "Musterstadt");
        plainData.put("DATUM", germanDate.format(datum));
        plainData.put("RECHNUNG", Integer.toString(rechnungsnummer));
        plainData.put("ANREDE", "Sehr geehrte Frau");
        plainData.put("AUF_DATUM", germanDate.format(auftragDatum));
        plainData.put("BETRAG", euroFormat.format(betrag));
        plainData.put("SKONTO_DATUM", germanDate.format(skontoDatum));
        plainData.put("SKONTO_BETRAG", euroFormat.format(skontobetrag));
        plainData.put("ZAHLUNG_DATUM", germanDate.format(zahlungDatum));
        plainData.put("AP_VORNAME", "Mister");
        plainData.put("AP_NAME", "Fantastic");
        
        List<Map<String,String>> lsts = new ArrayList<>();
        Map<String,String> lst;
        
        lst = new HashMap<>();
        lst.put("LST_TEXT", "Bemalte Eier");
        lst.put("LST_ZAHL", "2");
        lst.put("LST_EINZEL", euroFormat.format(3));
        lst.put("LST_SUM", euroFormat.format(6));
        lsts.add(lst);
        lst = new HashMap<>();
        lst.put("LST_TEXT", "Weiße Eier");
        lst.put("LST_ZAHL", "8");
        lst.put("LST_EINZEL", euroFormat.format(0.5));
        lst.put("LST_SUM", euroFormat.format(4));
        lsts.add(lst);
        lst = new HashMap<>();
        lst.put("LST_TEXT", "Schoko Eier (Kinder Edition)");
        lst.put("LST_ZAHL", "23");
        lst.put("LST_EINZEL", euroFormat.format(1.25));
        lst.put("LST_SUM", euroFormat.format(28.75));
        lsts.add(lst);
        
        // Initialize COM Subsystem -- this needs to be called on all threads
        // interacting with the COM objects!
        Ole32.INSTANCE.CoInitializeEx(Pointer.NULL, Ole32.COINIT_MULTITHREADED);
        // Initialize Factory for COM object creation
        Factory fact = new Factory();

        try {
            // The output will be written into an tempory file. The temp file
            // is created with the java utilites - the temp file is the deleted,
            // so word can write into it.
            //
            // The intention is not to be safe in the security sense, but safe in
            // the "this won't collide without bad intentions".
            File target = Files.createTempFile("output", ".pdf").toFile();
            target.delete();
            System.out.println("Output will be written to: " + target.getAbsolutePath());
            
            // Move resources from JAR into the filesystem, so that they are
            // accessible from word
            File tempFile1 = Files.createTempFile("GenerateDocument_", "docx").toFile();
            File tempFile2 = Files.createTempFile("signature", "png").toFile();
            
            copyResource("/eu/doppel_helix/jna/tlb/demo/Rechnung.docx", tempFile1);
            copyResource("/eu/doppel_helix/jna/tlb/demo/signature.png", tempFile2);
            File tempFile3 = writeQRCode(rechnungsnummer, datum, auftragDatum, skontoDatum, zahlungDatum, betrag, skontobetrag);
            
            // Start a new word instance
            Application wordApp = fact.createObject(Application.class);

            // Make word visible/invisible (invisible is default)
            wordApp.setVisible(false);

            // Open the source document 
            Document doc = wordApp.getDocuments().Open(tempFile1.getAbsolutePath(),
                    VARIANT_MISSING, VARIANT_MISSING, VARIANT_MISSING, VARIANT_MISSING, 
                    VARIANT_MISSING, VARIANT_MISSING, VARIANT_MISSING, VARIANT_MISSING, 
                    WdOpenFormat.wdOpenFormatAuto, VARIANT_MISSING, VARIANT_MISSING, VARIANT_MISSING,
                    VARIANT_MISSING, VARIANT_MISSING, VARIANT_MISSING);
            
            // Iterate alle the fields in the document - each field is checked:
            //  - is the field with plain string data in the plainData map
            //      then replace the result of the field with that data and
            //      unlink the field
            //  - is the field the special case "AP_UNTERSCHRIFT", then replace
            //      the contents of the field with the image of a signature
            //      this graphic could be dynamically created or fetched from DB
            //      in this demo it comes staticly from the JAR
            Fields fields = doc.getFields();
            for(int i = fields.getCount(); i > 0; i--) {
                Field f = fields.Item(i);
                String param = f.getResult().getText();
                String replacement = plainData.get(param);
                if(replacement != null) {
                    f.getResult().setText(replacement);
                    f.Unlink();
                } else if ("AP_UNTERSCHRIFT".equals(param)) {
                    Range targetRange = f.getResult();
                    targetRange.setText("");
                    InlineShape is = targetRange.getInlineShapes().AddPicture(
                            tempFile2.getAbsolutePath(),
                            Boolean.FALSE,
                            VARIANT_MISSING,
                            targetRange);
                    is.setLockAspectRatio(MsoTriState.msoTrue);
                    is.setWidth((Float) ( (float) (7 * CM_TO_POINT)));
                }
            }

            // Specal Handling for the tabular data. This is a possible approach
            // to be able to style a table. The alternative would be to create the
            // table totally dynamic.
            //
            // The Idea:
            // - Scan all tables
            // - foreach table check each row if it contains a field with the "LST_"
            //   prefix. 
            // - If so, the row is multiplied as often as necessary, to
            //   take all the data in the source array.
            // - replace the data in fields with the data from lsts-Map
            //
            // Negative side effect:
            // The clipboard is used to duplicate the row, so clipboard contents
            // is destroyed when running this.
            //
            // This is mighty ugly - better ideas are appretiated!
            Tables tabs = doc.getTables();
            for(int i = tabs.getCount(); i > 0; i--) {
                Table tab = tabs.Item(i);
                Rows rows = tab.getRows();
                INNER: for(int j = rows.getCount(); j > 0; j--) {
                    Row row = rows.Item(j);
                    if(row.getRange().getFields().getCount() > 0) {
                        for(int k = 0; k < (lsts.size() - 1); k++) {
                            row.getRange().Copy();
                            row.getRange().Paste();
                        }
                        for(int k = 0; k < lsts.size(); k++) {
                            row = rows.Item(j + k);
                            Map<String, String> data = lsts.get(k);
                            fields = row.getRange().getFields();
                            for (int l = fields.getCount(); l > 0; l--) {
                                Field f = fields.Item(l);
                                String param = f.getResult().getText();
                                String replacement = data.get(param);
                                if (replacement != null) {
                                    f.getResult().setText(replacement);
                                    f.Unlink();
                                }
                            }
                        }
                        break;
                    }
                }
            }
            
            // Iterate over the Inlineshaped to find the Image with the QRCODE
            // marker in the alternative TITLE
            Shapes shapes = doc.getShapes();
            for(int i = shapes.getCount(); i > 0; i--) {
                Shape s = shapes.Item(i);
                if(s.getType() == MsoShapeType.msoPicture && "QRCODE".equals(s.getAlternativeText().trim())) {
                    Shape s2 = shapes.AddPicture(
                            tempFile3.getAbsolutePath(),Boolean.FALSE,
                            VARIANT_MISSING, VARIANT_MISSING, VARIANT_MISSING,
                            VARIANT_MISSING, VARIANT_MISSING, VARIANT_MISSING);
                    Range a = s.getAnchor();
                    s2.getAnchor().SetRange(a.getStart(), a.getEnd());
                    s2.setRelativeHorizontalPosition(s.getRelativeHorizontalPosition());
                    s2.setRelativeHorizontalSize(s.getRelativeHorizontalSize());
                    s2.setRelativeVerticalPosition(s.getRelativeVerticalPosition());
                    s2.setRelativeVerticalSize(s.getRelativeVerticalSize());
                    s2.setLeft(s.getLeft());
                    s2.setTop(s.getTop());
                    s2.setWidth(s.getWidth());
                    s2.setHeight(s.getHeight());
                    s.Delete();
                }
            }
            
            // Export as PDF, special note to:
            // UseISO19005_1 if this is true, the embedded PNG file (the rabbit)
            // is replaced with a black square
            doc.ExportAsFixedFormat(
                    target.getAbsolutePath(), 
                    WdExportFormat.wdExportFormatPDF, 
                    Boolean.FALSE,
                    WdExportOptimizeFor.wdExportOptimizeForPrint,
                    WdExportRange.wdExportAllDocument,
                    0,
                    0,
                    WdExportItem.wdExportDocumentWithMarkup, 
                    Boolean.TRUE, 
                    Boolean.TRUE, 
                    WdExportCreateBookmarks.wdExportCreateNoBookmarks, 
                    Boolean.TRUE, 
                    Boolean.FALSE,
                    Boolean.FALSE, 
                    VARIANT_MISSING);
            
            // Shutdown the word instance
            // In this case saving is not wanted, the two further arguments
            // are to be ignored
            wordApp.Quit(WdSaveOptions.wdDoNotSaveChanges, VARIANT_MISSING, VARIANT_MISSING);
            
            // Remove temporary files
            tempFile1.delete();
            tempFile2.delete();
            tempFile3.delete();
        } finally {
            // Dispose the factory -- as there is no guarantee the finalizers are
            // run this should always be run!
            fact.disposeAll();
            // Uninitialize the COM subsystem - again: this needs to called on
            // all threads initialized by CoInitializeEx
            Ole32.INSTANCE.CoUninitialize();
        }
    }

    // Copy a named resource from inside the classpath to a real file
    private static void copyResource(String resource, File tempFile2) throws IOException {
        try(InputStream is = Automation_ReplaceFormField.class.getResourceAsStream(resource);
                OutputStream os = new FileOutputStream(tempFile2)) {
            byte[] buffer = new byte[20480];
            int read;
            while((read = is.read(buffer)) > 0) {
                os.write(buffer,0, read);
            }
        }
    }
    
    private static File writeQRCode(int rechnungsnummer, Date datum, Date auftragDatum, 
            Date skontoDatum, Date zahlungDatum, BigDecimal betrag, BigDecimal skontobetrag) throws IOException {
        SimpleDateFormat intFormat = new SimpleDateFormat("yyyy-MM-dd");
        File tempFile3 = Files.createTempFile("qrcode", "png").toFile();
        JsonObject ob = Json.createObjectBuilder()
                .add("rechnungsnummer", rechnungsnummer)
                .add("datum", intFormat.format(datum))
                .add("auftragDatum", intFormat.format(auftragDatum))
                .add("skontoDatum", intFormat.format(skontoDatum))
                .add("zahlungDatum", intFormat.format(zahlungDatum))
                .add("betrag", betrag)
                .add("skontobetrag", skontobetrag)
                .build();
        StringWriter sw = new StringWriter();
        Json.createWriter(sw).writeObject(ob);
        
        try {
            BitMatrix bm = new MultiFormatWriter().encode(sw.toString(), BarcodeFormat.QR_CODE, 300, 300);
            BufferedImage bi = MatrixToImageWriter.toBufferedImage(bm);
            ImageIO.write(bi, "png", tempFile3);
        } catch (WriterException ex) {
            throw new IOException("Failed to create QRCode Writer", ex);
        }
        
        return tempFile3;
    }
    
}
