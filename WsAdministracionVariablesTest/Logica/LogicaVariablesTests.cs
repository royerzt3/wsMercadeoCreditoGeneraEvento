using Microsoft.VisualStudio.TestTools.UnitTesting;
using WsAdministracionVariables.Logica;
using System;
using System.Collections.Generic;
using System.Text;

namespace WsAdministracionVariables.Logica.Tests
{
    [TestClass()]
    public class LogicaVariablesTests
    {

        private WsAdministracionVariables.Logica.LogicaVariables ws = null;

        [TestMethod()]
        public void LogicaVariablesTest()
        {
            
        }

        [TestMethod()]
        public void validacionCantidadArchivosTest()
        {
            ws = new WsAdministracionVariables.Logica.LogicaVariables();
            Assert.IsTrue(ws.ValidacionCantidadArchivos());
        }
    }
}