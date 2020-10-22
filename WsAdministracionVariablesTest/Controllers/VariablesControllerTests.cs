using Microsoft.VisualStudio.TestTools.UnitTesting;
using WsAdministracionVariables.Controllers;
using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Extensions.Configuration;

namespace WsAdministracionVariables.Controllers.Tests
{
    [TestClass()]
    public class VariablesControllerTests
    {

        private const string Expected = "Microsoft.AspNetCore.Mvc.OkObjectResult";
        private readonly IConfiguration _configuration = null;
        private WsAdministracionVariables.Controllers.VariablesController ws = null;

        [TestMethod()]
        public void VariablesControllerTest()
        {
            ws = new WsAdministracionVariables.Controllers.VariablesController(_configuration);
            var result = ws.Index().ToString().Trim();
            Assert.AreEqual(Expected, result);
        }

        [TestMethod()]
        public void IndexTest()
        {
            ws = new WsAdministracionVariables.Controllers.VariablesController(_configuration);
            Microsoft.AspNetCore.Mvc.IActionResult result = ws.Index();
            var okResult = result as Microsoft.AspNetCore.Mvc.OkObjectResult;
            Assert.AreEqual("{ saludo = Hola Mundo }", okResult.Value.ToString());
        }

        [TestMethod()]
        public void ExtraerVariablesTest()
        {
            
        }

        [TestMethod()]
        public void GuardarValoresTodosTest()
        {

        }
    }
}