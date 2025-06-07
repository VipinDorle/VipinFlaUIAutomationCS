using VipinFlaUIAutomationCS.Utility;

namespace VipinFlaUIAutomationCS.Tests
{
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\Resources\TestData1.xlsx");
            filePath = Path.GetFullPath(filePath);



            var td = ExcelUtility.ReadExcelFile(filePath, "TestSet1", "TC001");
            TestContext.WriteLine(td["Name"]);
        }

        [Test]
        public void Test1()
        {
            Assert.Pass();
        }
    }
}