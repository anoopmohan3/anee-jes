using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Whiteclass.TestCase
{
    [TestClass]
    public class Sample
    {

        public TestContext TestContext { get; set; }

        [TestInitialize()]
        public void Init()
        {
            Common.obj_st_GlobalTestSupport.sTestCaseName = TestContext.TestName;
            Common.InitializeSettings();
         }

        [TestMethod] //Test Case 101 
        public void LaunchForm()// with valid credentials
        {

            try
            {
                
                Methods.Methods.AppstLaunchSample();

                // Assert.IsNotNull(Methods.Methods.SOTAXHomeWindow);
                // Methods.Methods.SOTAXHomeWindow.Close();
                
               



            }
            catch (Exception e)
            {
                Common.Exception(e);
            }
        }

    }
}
