using System;
using NUnit.Framework;
using FaresListImplementation;

namespace FaresListImplementation.Tests
{
    [TestFixture]
    public class FileTests
    {
        [Test]
        public void TestFileExists()
        {
           var fileFunction = new FileFunctions();
            string fileName = @"c:\temp\FaresListImplementation.xlsx";
            Assert.IsTrue(fileFunction.CheckFileExists(fileName));
        }
    }
}
