using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyComments
{
    class Program
    {
        static void Main(string[] args)
        {
            SpreadsheetDocument docWithComments = SpreadsheetDocument.Open("C:\\Temp\\TestComment2.xlsx", true);
            SpreadsheetDocument docWithoutComments = SpreadsheetDocument.Open("C:\\Temp\\TestComment2Target.xlsx", true);
            WorkbookPart wbpWithComments = docWithComments.WorkbookPart;
            WorkbookPart wbpWithoutComments = docWithoutComments.WorkbookPart;

            for (int i = 0; i < wbpWithComments.WorksheetParts.Count(); i++)
            {
                if (wbpWithoutComments.WorksheetParts.ElementAt(i).WorksheetCommentsPart == null)
                {
                    wbpWithoutComments.WorksheetParts.ElementAt(i).AddNewPart<WorksheetCommentsPart>();
                    wbpWithoutComments.WorksheetParts.ElementAt(i).WorksheetCommentsPart.Comments = new Comments();
                    wbpWithoutComments.WorksheetParts.ElementAt(i).WorksheetCommentsPart.Comments.CommentList = new CommentList();
                }
                for (int j = 0; j < wbpWithComments.WorksheetParts.ElementAt(i).WorksheetCommentsPart.Comments.Count(); j++)
                {
                    var a = wbpWithComments.WorksheetParts.ElementAt(i).WorksheetCommentsPart.Comments.CommentList.ElementAt(j);
                    wbpWithoutComments.WorksheetParts.ElementAt(i).WorksheetCommentsPart.Comments.CommentList.Append(a.CloneNode(true));
                    var b = wbpWithoutComments.WorksheetParts.ElementAt(i).WorksheetCommentsPart.Comments.CommentList.ElementAt(j);
                }
            }

        }
    }
}
