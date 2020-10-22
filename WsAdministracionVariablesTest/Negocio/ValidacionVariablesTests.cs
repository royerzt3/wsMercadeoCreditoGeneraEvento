using Microsoft.VisualStudio.TestTools.UnitTesting;
using WsAdministracionVariables.Negocio;
using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Extensions.Configuration;

namespace WsAdministracionVariables.Negocio.Tests
{
    [TestClass()]
    public class ValidacionVariablesTests
    {

        private readonly IConfiguration _configuration = null;
        private WsAdministracionVariables.Negocio.ValidacionVariables ws = null;

        [TestMethod()]
        public void ValidacionVariablesTest()
        {
            ws = new WsAdministracionVariables.Negocio.ValidacionVariables(_configuration);
            //string uri = ws.uriValidacionVariables;
            //Console.WriteLine(uri);
        }

        [TestMethod()]
        public void construccionObjetoPeticionTest()
        {

        }

        [TestMethod()]
        public void realizarPeticionValidacionTest()
        {

        }

        [TestMethod()]
        public void validacionValoresVariablesTest()
        {

        }
    }
}