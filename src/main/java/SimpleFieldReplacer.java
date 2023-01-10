import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSimpleField;

import java.util.ArrayList;
import java.util.List;

/**
 * Replaces all SimpleFields in a paragraph with the pure text each field contains.
 * After usage, expected is that the paragraph will look the same for a viewer using Word.
 *
 * Some tech:
 * In POI, the XML based Word doxument is represented by XWPFDocument.
 * Each paragraph in the doc is a XWPFParagraph.
 * Each XWPFParagraph is further divided into "runs", the class XWPFRun. Each run holds formatting stuff and more for a bit of text.
 *
 * A simple field is a CTSimpleField, often (always?) placed inside a paragraph.
 * The field holds a run, and that run holds the actual text that is shown.
 *
 * The XML would look something like this:
 *
 * <pre>{@code
 * <w:fldSimple w:instr=" DOCPROPERTY field_name \* MERGEFORMAT ">
 *   <w:r>
 *     <w:t>The text shown to the user< /w:t>
 *   </w:r>
 * </w:fldSimple>
 * }</pre>
 *
 * @see <a href="http://officeopenxml.com/WPfields.php">Office XML Open: Wordprocessing Fields</a>
 *
 */
public class SimpleFieldReplacer {

    /**
     * Replaces all simple fields in the paragraph with their pure text.
     * Works like:
     * 1. create a list of each simple field's run in the paragraph.
     * 2. For each run in the list:
     *  a. find its index of all runs in the paragrah.
     *  b. copy the run to a new run on the same index (I guess pushing the simple field with its run to index+1, but that doesn't matter)
     * 3. Find all simple fields and remove them, thereby also removing each field's run.
     *
     * Note the comment in #replaceSimpleFieldsWithTextFromInsideRun about a potential bug in POI that must be fended of with a magic number.
     */
    public XWPFParagraph inlineReplaceSimpleFieldsWithText(XWPFParagraph xwpfParagraph) {
        var simpleFieldAndRunsHolder  = findRunForEachSimpleFields(xwpfParagraph);
        return replaceSimpleFieldsWithTextFromInsideRun(xwpfParagraph, simpleFieldAndRunsHolder);
    }

    private List<XWPFRun> findRunForEachSimpleFields(XWPFParagraph xwpfParagraph) {

        var runs  = new ArrayList<XWPFRun>();

        for (CTSimpleField simpleFieldToRemove : xwpfParagraph.getCTP().getFldSimpleArray()) {

            for (XWPFRun run : xwpfParagraph.getRuns()) {
                if (run.getCTR().getDomNode().getParentNode().getNodeName().equals("w:fldSimple")){
                    var candidateFieldNode = run.getCTR().getDomNode().getParentNode();
                    if(candidateFieldNode.equals(simpleFieldToRemove.getDomNode())){
                        //System.out.println("We have a match! " + run.getText(0));
                        runs.add(run);
                    }
                }
            }
        }
        return runs;
    }

    private XWPFParagraph replaceSimpleFieldsWithTextFromInsideRun(XWPFParagraph xwpfParagraph, List<XWPFRun> runs) {

        //Add new run with same text as the one inside the SimpleField
        for(int i=0; i < runs.size(); i++) {
            var oldRun = runs.get(i);
            int runIndex = findRunIndexInParagraph(xwpfParagraph, oldRun);
            // NOTE: The minus i below is needed because xwpfParagraph doesn't seem to update its internal
            // count of how many Runs it contains. So the pos passed must be what the pos should be before any
            // previous calls to insertNewRun. This old bug seems to be something similar: https://bz.apache.org/bugzilla/show_bug.cgi?id=57829
            // (Or perhaps this is some kind of bug on my side ¯\_(ツ)_/¯ )
            XWPFRun newRun = xwpfParagraph.insertNewRun(runIndex - i);
            copyEverythingFromOldRunToNew(oldRun, newRun);
        }

        // Remove all SimpleFields
        int nbrOfFieldsToRemove = xwpfParagraph.getCTP().getFldSimpleArray().length;
        for(int i = 0; i < nbrOfFieldsToRemove; i++){
            xwpfParagraph.getCTP().removeFldSimple(0); // we always remove at index zero, because the removal makes the list smaller each iteration
        }
        return xwpfParagraph;
    }

    private void printRuns(XWPFParagraph xwpfParagraph) {
        xwpfParagraph.getRuns().forEach(System.out::println);
    }

    private int findRunIndexInParagraph(XWPFParagraph xwpfParagraph, XWPFRun run) {
            for(int i =0; i < xwpfParagraph.getRuns().size(); i++){
           if(xwpfParagraph.getRuns().get(i).equals(run)){
               return i;
           }
        }
        throw new RuntimeException("Couldn't find expected Run in paragraph. " +
                "Run text: '" + run.getText(0) +"' Paragraph text '" +xwpfParagraph.getText() + "' ");
    }

    private void copyEverythingFromOldRunToNew(XWPFRun oldRun, XWPFRun newRun) {
        newRun.setText(oldRun.getText(0));

        // strange, seems like it's -2 sometimes which makes the text invisible.
        if(oldRun.getFontSize()>0) {
            newRun.setFontSize(oldRun.getFontSize());
        }

        newRun.setBold(oldRun.isBold());
        newRun.setItalic(oldRun.isItalic());
        newRun.setUnderline(oldRun.getUnderline());
        newRun.setColor(oldRun.getColor());
        newRun.setFontFamily(oldRun.getFontFamily());
        newRun.setSubscript(oldRun.getSubscript());

        // These are commented out as they exist in Poi v. 3.15, but not in 3.10_FINAL
//        newRun.setCapitalized(oldRun.isCapitalized());
//        newRun.setCharacterSpacing(oldRun.getCharacterSpacing());
//        newRun.setStrikeThrough(oldRun.isStrikeThrough());
//        newRun.setDoubleStrikethrough(oldRun.isDoubleStrikeThrough());
//        newRun.setEmbossed(oldRun.isEmbossed());
//        newRun.setImprinted(oldRun.isImprinted());
//        newRun.setKerning(oldRun.getKerning());
//        newRun.setSmallCaps(oldRun.isSmallCaps());
//        newRun.setShadow(oldRun.isShadowed());

    }
}
