import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Test;
import org.openxmlformats.schemas.officeDocument.x2006.customProperties.CTProperty;

import java.io.*;

import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.MatcherAssert.assertThat;

public class DocPropTest {

    @Test
    public void sdfsdfsdf()throws IOException {

        String copyFilePath= createCopyOf("testpropdoc.docx");
        XWPFDocument doc = new XWPFDocument(new FileInputStream(copyFilePath));


        var userProps = doc.getProperties().getCustomProperties().getUnderlyingProperties();
        for (CTProperty ctProp : userProps.getPropertyList()) {
            System.out.println(ctProp.getName() + ":" + ctProp.getLpwstr());
        }

        userProps.getPropertyList().get(0).setLpwstr("Hello");
        doc.enforceUpdateFields();

        saveAndClose(doc, copyFilePath);

    }

    private String createCopyOf(String filename) throws IOException {
        String path = getClass().getResource(filename).getPath();
        String copyPath = path.substring(0,path.lastIndexOf(".") )+ "_copy.docx";

        FileInputStream fileInputStream = new FileInputStream(path);
        XWPFDocument origDoc = new XWPFDocument(fileInputStream);
        fileInputStream.close();

        FileOutputStream copyOut = new FileOutputStream(copyPath);
        origDoc.write(copyOut);
        copyOut.close();

        return copyPath;
    }

    private void saveAndClose(XWPFDocument doc, String filePath) throws IOException {
        var out = new FileOutputStream(filePath);
        doc.write(out);
        out.close();
    }

}
