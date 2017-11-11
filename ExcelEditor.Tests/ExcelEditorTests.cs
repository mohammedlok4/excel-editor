using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace ExcelEditor.Tests
{
    [TestClass]
    public class ExcelEditorTests
    {
        [TestMethod]
        public void GetWorkSheetCountTest()
        {
            List<string> sheets = ExcelEditorServices.GetWorksheetNames("Chart.xlsx");
            int count = sheets.Count;
            Assert.AreEqual(4, count);
        }

        [TestMethod]
        public void RemoveFormulasTest()
        {
            bool success = ExcelEditorServices.ReplaceFormulasWithValues("Chart.xlsx");
            Assert.AreEqual(success, true);
        }

        [TestMethod]
        public void GetTextBoxesCountFromSheet1()
        {
            List<string> textBoxes = ExcelEditorServices.GetTextBoxes("Chart.xlsx", "Sheet1");
            int count = textBoxes.Count;
            Assert.AreEqual(count, 1);
        }

        [TestMethod]
        [ExpectedException(typeof(NullReferenceException))]
        public void GetTextBoxesCountFromSheet3 ()
        {
            List<string> textBoxes = ExcelEditorServices.GetTextBoxes("Chart.xlsx", "Sheet3");
        }
    }

    
}
