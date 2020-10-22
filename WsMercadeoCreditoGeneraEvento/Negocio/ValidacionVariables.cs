using BibliotecaSimulador.Pojos;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using static BibliotecaSimulador.Negocio.Plazos;

namespace WsAdministracionVariables.Negocio
{
    public class ValidacionVariables
    {
        public BibliotecaSimulador.Pojos.TipoProductoPeticion tipoProductoPeticion;
        public List<TipoProductoFamilia> tipoProductoFamilia { get; set; }

        public ValidacionVariables(IConfiguration configuration)
        {
        }
        public void construccionObjetoPeticion(List<TipoProductoFamilia> tipoProductoFamilias)
        {
            try
            {
                tipoProductoPeticion = new TipoProductoPeticion();
                this.tipoProductoFamilia = tipoProductoFamilias;
                tipoProductoPeticion.Familias = tipoProductoFamilias
                    .Where(e => e.fiFamiliaId > 0)
                    .Select((f, index) => new FamiliaPeticion
                    {
                        IdProducto = f.fiTipoProductoId,
                        DescFamilia = f.fcDescripcionFamilia
                        ,
                        IdFamilia = f.fiFamiliaId
                        ,
                        Variables = f.variablesCabeceras
                        .GroupBy(g => g.fiVariableId)
                        .Where(i => i.Key > 0)
                        .Select(s => s.Key.ToString()).ToList()
                        ,
                        Plazos = f.variablesCabeceras
                        .Select(j=>j.variables)
                        .Select((m,indiceVariables)=>m[indiceVariables])
                        .GroupBy(k=>k.plazo)
                        .Select(n=>n.Key.ToString()).ToList()
                    }).ToList();
            }
            catch (Exception e)
            {
                this.tipoProductoPeticion = null;
                throw new BibliotecaSimulador.DefiniedExceptions.DefiniedGenericException(e.Message, e);
            }
        }
        public ObjetoRespuestaVariables RealizarPeticionValidacion()
        {
            try
            {

                ConsultaVariables consultaVariables = new ConsultaVariables();
                EntidadesPlazos.EntPeticionesMultiples ent = new EntidadesPlazos.EntPeticionesMultiples();
                ent.Familias = this.tipoProductoPeticion.Familias.Select(
                    e=>new EntidadesPlazos.EntListProd 
                    { 
                        DescFamilia = e.DescFamilia,
                        IdFamilia = e.IdFamilia,
                        IdProducto = e.IdProducto,
                        Plazos = e.Plazos,
                        Variables = e.Variables
                    }).ToList();
                var plazosValidos = consultaVariables.ConsultaLimPor4Va2daOpcion(ent);
                if (!(plazosValidos is null))
                {
                    ObjetoRespuestaVariables objetoRespuesta = JsonConvert.DeserializeObject<ObjetoRespuestaVariables>(JsonConvert.SerializeObject(plazosValidos));
                    return objetoRespuesta;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception e)
            {
                return null;
                throw new BibliotecaSimulador.DefiniedExceptions.DefiniedGenericException(e.Message, e);
            }
        }

        public bool ValidacionValoresVariables(ObjetoRespuestaVariables objetoRespuesta)
        {
            try
            {
                if (!objetoRespuesta.Error && objetoRespuesta.Mensaje == "Exito")
                {
                    if (objetoRespuesta.Valores.Count > 0)
                    {
                        foreach (var valorRespuesta in objetoRespuesta.Valores)
                        {
                            foreach (var familia in this.tipoProductoFamilia.Where(c => 
                            c.fiFamiliaId == int.Parse(valorRespuesta.IdFamilia) && 
                            c.fiTipoProductoId == int.Parse(valorRespuesta.IdProducto)))
                            {
                                foreach (var variableBRMS in valorRespuesta.ListaPLazos)
                                {
                                    foreach (var variable in familia.variablesCabeceras.Where(v => v.fiVariableId == int.Parse(variableBRMS.IDVariable)))
                                    {
                                        foreach (var valorBRMS in variableBRMS.Variables)
                                        {
                                            foreach (var valor in variable.variables.Where(vb => vb.plazo == ushort.Parse(valorBRMS.IdPlazo)))
                                            {
                                                int indiceFamilia = this.tipoProductoFamilia.IndexOf(familia);
                                                int indiceVariableCabecera = this.tipoProductoFamilia[indiceFamilia].variablesCabeceras.IndexOf(variable);
                                                int indiceVariable = this.tipoProductoFamilia[indiceFamilia].variablesCabeceras[indiceVariableCabecera].variables.IndexOf(valor);
                                                decimal valorVariable = this.tipoProductoFamilia[indiceFamilia].variablesCabeceras[indiceVariableCabecera].variables[indiceVariable].valor;
                                                decimal valorMinimo = decimal.Parse(valorBRMS.Valor_Minimo);
                                                decimal valorMaximo = decimal.Parse(valorBRMS.Valor_Maximo);
                                                if (valorVariable >= valorMinimo && valorVariable <= valorMaximo)
                                                {
                                                    this.tipoProductoFamilia[indiceFamilia].variablesCabeceras[indiceVariableCabecera].variables[indiceVariable].encontrado = true;
                                                    this.tipoProductoFamilia[indiceFamilia].variablesCabeceras[indiceVariableCabecera].variables[indiceVariable].esValido = true;
                                                }
                                                else
                                                {
                                                    this.tipoProductoFamilia[indiceFamilia].variablesCabeceras[indiceVariableCabecera].variables[indiceVariable].encontrado = true;
                                                    this.tipoProductoFamilia[indiceFamilia].variablesCabeceras[indiceVariableCabecera].variables[indiceVariable].esValido = false;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (Exception e)
            {
                return false;
                throw new BibliotecaSimulador.DefiniedExceptions.DefiniedGenericException(e.Message, e);
            }
        }
    }
}
