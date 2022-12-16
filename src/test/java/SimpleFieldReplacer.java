import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSimpleField;

import java.util.HashMap;
import java.util.Map;

/**
 * Replaces all SimpleFields in a paragraph with the pure text each field contains.
 * After usage, expected is that the paragraph will look the same for a viewer using Word.
 */
public class SimpleFieldReplacer {

    private XWPFParagraph xwpfParagraph;

    public SimpleFieldReplacer(XWPFParagraph xwpfParagraph) {
        this.xwpfParagraph = xwpfParagraph;
    }

    public void inlineReplaceSimpleFieldsWithText() {
        var simpleFieldAndRunsHolder  = findRunForEachSimpleFields();
        replaceSimpleFieldsWithTextFromInsideRun(simpleFieldAndRunsHolder);
    }

    private HashMap<CTSimpleField, XWPFRun> findRunForEachSimpleFields() {

        var simpleFieldAndRunsHolder  = new HashMap<CTSimpleField, XWPFRun>();

        for (CTSimpleField simpleFieldToRemove : xwpfParagraph.getCTP().getFldSimpleArray()) {

            for (XWPFRun run : xwpfParagraph.getRuns()) {
                if (run.getCTR().getDomNode().getParentNode().getNodeName().equals("w:fldSimple")){
                    var candidateFieldNode = run.getCTR().getDomNode().getParentNode();
                    if(candidateFieldNode.equals(simpleFieldToRemove.getDomNode())){
                        //System.out.println("We have a match! " + run.getText(0));

                        if(simpleFieldAndRunsHolder.containsKey(simpleFieldToRemove)){
                            throw new RuntimeException("Until proven otherwise, a SimpleField should only contain a single Run. It now seems I've been proven otherwise");
                        }
                        simpleFieldAndRunsHolder.put(simpleFieldToRemove, run);
                    }
                }
            }
        }
        return simpleFieldAndRunsHolder;
    }

    private void replaceSimpleFieldsWithTextFromInsideRun(HashMap<CTSimpleField, XWPFRun> simpleFieldAndRunHolder) {

        //Add new run with same text as the one inside the SimpleField
        for (Map.Entry<CTSimpleField, XWPFRun> entry : simpleFieldAndRunHolder.entrySet()) {

            int runIndex = findRunIndexInParagraph(entry.getValue());
            XWPFRun newRun = xwpfParagraph.insertNewRun(runIndex);
            copyEverythingFromOldRunToNew(entry.getValue(), newRun);
        }

        // Remove all SimpleFields
        int nbrOfFieldsToRemove = xwpfParagraph.getCTP().getFldSimpleArray().length;
        for(int fieldindex = 0; fieldindex<nbrOfFieldsToRemove; fieldindex++){
            xwpfParagraph.getCTP().removeFldSimple(fieldindex);
        }
    }

    private int findRunIndexInParagraph(XWPFRun run) {
        for(int i =0; i < xwpfParagraph.getRuns().size(); i++){
           if(xwpfParagraph.getRuns().get(i).equals(run)){
               return i;
           }
        }
        throw new RuntimeException("Couldn't find expected Run i paragraph. " +
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
