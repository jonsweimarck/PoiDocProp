import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.MatcherAssert.assertThat;

// Makes use of the test doc testpropdoc.docx in the main/resources. The doc is copied to testpropdoc_copy.docx in the
// test, so the original doc is never altered. The testpropdoc_copy.docx can be opened in Word to see the result of the test.
public class DocPropTest {

    private String copyFilePath;
    private XWPFDocument doc;

    @Test
    public void replaceSingleSimpleFieldInEachParagraph() throws IOException {

        copyFilePath= createCopyOf("testpropdoc.docx");
        doc = new XWPFDocument(new FileInputStream(copyFilePath));

        // We start with a doc with some paragraphs, where two of them contain one simple field each.
        // sanity check:
        String expectedFullText = "Några fältTestprop1: <<testprop1>>Testprop2: <<testprop2>>Testprop3: <<testprop3>>Testprop4: <<testprop4>>";
        assertThat(getNumberOfParagraphs(doc), is(8));
        assertThat(getNumerOfSmartFields(doc), is (2));
        assertThat(getAllText(doc), is(expectedFullText));


        var paras= doc.getParagraphs();
        for(XWPFParagraph para : paras) {
            var simpleFieldReplacer = new SimpleFieldReplacer(para);
            simpleFieldReplacer.inlineReplaceSimpleFieldsWithText();
        }

        saveAndClose(doc, copyFilePath);

        // Now we want the same number of paragraphs and the same document text, but zero smart fields
        doc = new XWPFDocument(new FileInputStream(copyFilePath));
        assertThat(getNumberOfParagraphs(doc), is(8));
        assertThat(getNumerOfSmartFields(doc), is (0));
        assertThat(getAllText(doc), is(expectedFullText));

    }

    private String getAllText(XWPFDocument doc) {
        String returnString = "";

        for (XWPFParagraph para : doc.getParagraphs()) {
            returnString = returnString + para.getText();
        }

        return returnString;
    }

    private int getNumerOfSmartFields(XWPFDocument doc) {
        int nbr = 0;
        for (XWPFParagraph para : doc.getParagraphs()) {
            nbr += para.getCTP().getFldSimpleArray().length;
        }
        return nbr;
    }

    private int getNumberOfParagraphs(XWPFDocument doc) {
        return doc.getParagraphs().size();
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
