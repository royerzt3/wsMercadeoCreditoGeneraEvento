using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using BibliotecaSimulador.LecturaVariablesSimulador;
using BibliotecaSimulador.Pojos;
using BibliotecaSimulador.SimuladorDAO;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using WsAdministracionVariables.Negocio;

namespace WsAdministracionVariables.Controllers
{
    [ApiController]
    [Route("[controller]/api/")]
    public class GeneraEventoController : ControllerBase
    {
        private BibliotecaSimulador.Logs.Logg _log;
        private readonly IConfiguration _configuration;
        private readonly string[] extensionesPermitidas = { ".XLS", ".XLSX" };
        private string Nombrelog { get; set; }
        public GeneraEventoController(IConfiguration configuration)
        {
            this._configuration = configuration;
            this.Nombrelog = this._configuration.GetValue<string>("NombreLogVariables");
            this._log = new BibliotecaSimulador.Logs.Logg(this.Nombrelog);
        }
        public IActionResult Index()
        {
            return Ok(new { saludo = "Hola Mundo" });
        }
        [Produces("application/json")]
        [HttpPost("extraerVariablesM/evento/{id}/")]
        public async Task<IActionResult> ExtraerVariables(int id, int? pais)
        {
            var rutaArchivo = "";
            try
            {
                Stopwatch time = new Stopwatch();
                time.Start();
                this._log.WriteInfo($"Inicia el servicio para extraer las variables");
                List<TipoProductoFamilia> tipoProductoFamilias = null;
                int cantidadArchivos = Request.Form.Files.Count;

                if (cantidadArchivos > 0 && cantidadArchivos == 1)
                {
                    var archivoExcelPeticion = Request.Form.Files.Select(e => e).ToList()[0];
                    string extensionArchivo = Path.GetExtension(archivoExcelPeticion.FileName).ToUpperInvariant();

                    if (string.IsNullOrEmpty(extensionArchivo) || this.extensionesPermitidas.Contains(extensionArchivo))
                    {
                        ArchivoExcelEventos archivoExcel = null;
                        rutaArchivo = $"{Path.GetTempPath()}variablesEvento{DateTime.Now.Ticks}.xlsx";
                        using (var stream = new FileStream(rutaArchivo, FileMode.Create))
                        {
                            await archivoExcelPeticion.CopyToAsync(stream);
                            stream.Close();
                            archivoExcel = new ArchivoExcelEventos(rutaArchivo);//Se envia la ruta temporal del archivo e:\Users\ppastrana\AppData\Local\Temp\variables637388951388493557.xlsx 
                        }
                        archivoExcel.RutaTemportal = rutaArchivo;
                        tipoProductoFamilias = archivoExcel.ExcelUsingEPPlus().ToList();
                        time.Stop();
                        this._log.WriteAndCountService($"Método:{nameof(ExtraerVariables)}-> Se ejecuto correctamente",
                            new Dictionary<string, int>
                            {
                                {
                                    nameof(WsAdministracionVariables),
                                    Convert.ToInt32(time.ElapsedMilliseconds)
                                }
                            });
                        return Ok(
                            new RespuestaOK
                            {
                                respuesta =
                                new
                                {
                                    productosFamilias = tipoProductoFamilias,
                                    esValido = !(tipoProductoFamilias.Any(e => e.fiFamiliaId == 0))
                                }
                            });
                    }
                    else
                    {
                        time.Stop();
                        this._log.WriteAndCountService($"Método:{nameof(ExtraerVariables)}-> No es un archivo el archivo que se cargo",
                            new Dictionary<string, int>
                            {
                                {
                                    nameof(WsAdministracionVariables),
                                    Convert.ToInt32(time.ElapsedMilliseconds)
                                }
                            });
                        return BadRequest(
                            new RespuestaError400
                            {
                                errorInfo = string.Empty,
                                errorMessage = $"No es un archivo de Excel"
                            });
                    }
                }
                else
                {
                    time.Stop();
                    this._log.WriteAndCountService($"Método:{nameof(ExtraerVariables)}-> Se cargo más de un archivo o no se cargo ningún archivo",
                        new Dictionary<string, int>
                        {
                            {
                                nameof(WsAdministracionVariables),
                                Convert.ToInt32(time.ElapsedMilliseconds)
                            }
                        });
                    return BadRequest(
                        new RespuestaError400
                        {
                            errorInfo = string.Empty,
                            errorMessage = cantidadArchivos > 0 ? "Solo se puede subir un archivo" : "No subiste ningún archivo"
                        });
                }
            }
            catch (Exception ex)
            {
                if (ex is UnauthorizedAccessException)
                {
                    if (this._log is null)
                    {
                        this._log = new BibliotecaSimulador.Logs.Logg(this.Nombrelog);
                        this._log.WriteErrorService(ex, nameof(WsAdministracionVariables));
                        return StatusCode(StatusCodes.Status500InternalServerError,
                            new RespuestaError
                            {
                                errorMessage = this._configuration.GetValue<string>("Mensajes:Errores:CrearArchivoExcel")
                            });
                    }
                    this._log.WriteErrorService(ex, nameof(WsAdministracionVariables));
                    return StatusCode(StatusCodes.Status500InternalServerError,
                        new RespuestaError
                        {
                            errorMessage = this._configuration.GetValue<string>($"Mensajes:Errores:Generico")
                        });
                }
                else if (ex is ArgumentNullException)
                {
                    if (this._log is null)
                    {
                        this._log = new BibliotecaSimulador.Logs.Logg(this.Nombrelog);
                        this._log.WriteErrorService(ex, nameof(WsAdministracionVariables));
                        return StatusCode(StatusCodes.Status500InternalServerError,
                            new RespuestaError
                            {
                                errorMessage = this._configuration.GetValue<string>("Mensajes:Errores:Argumento")
                            });
                    }
                    this._log.WriteErrorService(ex, nameof(WsAdministracionVariables));
                    return StatusCode(StatusCodes.Status500InternalServerError,
                        new RespuestaError
                        {
                            errorMessage = this._configuration.GetValue<string>("Mensajes:Errores:Generico")
                        });
                }
                else
                {
                    this._log.WriteErrorService(ex, nameof(WsAdministracionVariables));
                    return StatusCode(StatusCodes.Status500InternalServerError,
                        new RespuestaError
                        {
                            errorMessage = this._configuration.GetValue<string>("Mensajes:Errores:Generico")
                        });
                }
            }
            finally
            {
                if (System.IO.File.Exists(rutaArchivo))
                {
                    System.IO.File.Delete(rutaArchivo);
                }
            }
        }
        [HttpPost("guardarVariablesTodo/usuario/{id}")]
       


        [HttpGet("Test")]
        public string Credits()
        {
            Dictionary<string, string> creditos = new Dictionary<string, string>();

            creditos.Add("Tipo proyecto", "Servicio Web");
            creditos.Add("Nombre", "WSVariables");
            creditos.Add("Version Net Core", "3.1");
            creditos.Add("Area", "Credito");
            creditos.Add("Servidor Activo", "OK");
            creditos.Add("Version", "3.1");
            creditos.Add("Cambio", "Metodos de Prueba con respuesta Json");
            return System.Text.Json.JsonSerializer.Serialize(creditos);
        }

    }
}
