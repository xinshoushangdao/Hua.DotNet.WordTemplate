using NPOI.XWPF.UserModel;

namespace Hua.DotNet.WordTemplate
{
    public static class NPOI
    {
        public static XWPFDocument Clone(this XWPFDocument srcDocx)
        {
            var newDocx  = new XWPFDocument();
            //#region 编号相关

            //#endregion

            return newDocx;
        }
        public static XWPFParagraph Clone(this XWPFParagraph srcPara, XWPFDocument? descDocx = null, bool cloneText = false)
        {
            XWPFParagraph output =null;
            if (descDocx != null)
            {
                if (cloneText)
                    return new XWPFParagraph(srcPara.GetCTP(), descDocx.CreateParagraph().Body);

                output =descDocx.CreateParagraph();
            }
            else
            {
                output = srcPara.Document.CreateParagraph();
            }

            // 设置新段落的属性
            output.Alignment = srcPara.Alignment;
            output.IndentationFirstLine = srcPara.IndentationFirstLine;
            output.IndentationLeft = srcPara.IndentationLeft;
            output.IndentationRight = srcPara.IndentationRight;
            //output.Style = srcDocx.Style;
            //output.SpacingAfter = srcDocx.SpacingAfter;
            //output.SpacingBefore = srcDocx.SpacingBefore;
            output.SetNumID(srcPara.GetNumID());
            if (!cloneText) return output;
            foreach (var src in srcPara.Runs)
            {
                var newRun = output.CreateRun();
                newRun.FontSize = src.FontSize;
                newRun.FontFamily = src.FontFamily;
                newRun.IsBold = src.IsBold;
                newRun.IsItalic = src.IsItalic;
                newRun.IsStrikeThrough = src.IsStrikeThrough;
                newRun.IsDoubleStrikeThrough = src.IsDoubleStrikeThrough;
                newRun.IsSmallCaps = src.IsSmallCaps;
                newRun.IsCapitalized = src.IsCapitalized;
                newRun.IsShadowed = src.IsShadowed;
                newRun.CharacterSpacing = src.CharacterSpacing;
                newRun.Kerning = src.Kerning;
                newRun.SetColor(src.GetColor());
                newRun.Underline = src.Underline;
                newRun.SetColor(src.GetColor());
                newRun.SetText(src.Text);
            }
            return output;
        }

        //public static XWPFParagraph Create(this XWPFParagraph srcPara, string paraContent, XWPFDocument? descDocx = null)
        //{
        //    var newPara = srcPara.Clone(descDocx, true);
        //    return newPara;
        //}

        public static void Create(this XWPFDocument srcDocx, XWPFDocument descDocx, int srcIndex, int descIndex, string placeholder, string value)
        {
            var para = srcDocx.Paragraphs[srcIndex];
            var descPara = para.Clone(descDocx, true);
            if (!string.IsNullOrEmpty(placeholder))
            {
                descPara.ReplaceText(placeholder, value);
            }

            //if (descDocx.Paragraphs.Count <= descIndex) return;
            descDocx.SetParagraph(descPara, descIndex);
        }

    }
}
