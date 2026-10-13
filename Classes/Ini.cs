using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Lotes_Hornitos.Classes
{

    internal class Ini
    {

        #region variables

        /// <summary>
        /// Variable local sap compañia
        /// </summary>
        private static SAPbobsCOM.Company oCompany;

        /// <summary>
        /// Variable local sap application
        /// </summary>
        private static SAPbouiCOM.Application oApp;

        /// <summary>
        /// Variable local ensamble 
        /// </summary>
        public static Assembly asm;
        private FormLotesEntrada formLotesEntrada;
        private FormLotesSalida formLotesSalida;
        private FormLotesConteo formLotesConteo;
        public static string Tipo;



        #endregion

        #region Constructor

        private Ini()
        {
            conecctSap();
            //LoadMenu();
            Filters();
            Events();
        }

        /// <summary>
        /// Metodo inicio
        /// </summary>
        /// <returns>Retorna Bool</returns>
        internal static bool Run()
        {
            try
            {
                _ = new Ini();
                oApp.SetStatusBarMessage("Add-On " + asm.GetName().Name + " conectado exitosamente.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                return true;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error Add-On " + asm.GetName().Name + " " + ex.Message.ToString() + " " + ex.StackTrace.ToString());
                return false;
            }
        }

        #endregion


        #region Eventos locales de inicio

        /// <summary>
        /// Conectar sesion de sap
        /// </summary>
        private void conecctSap()
        {
            SAPbouiCOM.SboGuiApi oGuiApi = null;
            string strConnect = null;
            asm = GetType().Assembly;
            if (!Environment.GetCommandLineArgs().GetValue(1).Equals(""))
            {
                if (!Environment.GetCommandLineArgs()[1].Equals(""))
                {
                    strConnect = Environment.GetCommandLineArgs()[1].ToString();
                    oGuiApi = new SAPbouiCOM.SboGuiApi();
                    oGuiApi.Connect(strConnect);
                    oApp = oGuiApi.GetApplication();
                    oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();

                }

            }
            else
            {
                strConnect = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
                oGuiApi = new SAPbouiCOM.SboGuiApi();
                oGuiApi.Connect(strConnect);
                oApp = oGuiApi.GetApplication();
                oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();

            }
        }

        /// <summary>
        /// Metodo para agregar menus nuevos
        /// </summary>
        private void LoadMenu()
        {
            int countMenu = 0;
            string sPath = null, stringMenu = string.Empty, strdoc = string.Empty;
            SAPbouiCOM.Menus SboMenu = null;
            SAPbouiCOM.MenuItem oMenuItem = null;
            SAPbouiCOM.MenuCreationParams oMenuParams = null;
            SAPbouiCOM.Form oFormMenu = null;
            try
            {
                oMenuParams = (SAPbouiCOM.MenuCreationParams)oApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

                sPath = System.Windows.Forms.Application.StartupPath;

                oMenuItem = oApp.Menus.Item("43535");
                SboMenu = oMenuItem.SubMenus;
                countMenu = SboMenu.Count;
                if (SboMenu.Exists("MnuRecAut"))
                {
                    SboMenu.RemoveEx("MnuRecAut");
                }
                oFormMenu = oApp.Forms.GetFormByTypeAndCount(169, 1);
                oFormMenu.Freeze(true);
                oMenuParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oMenuParams.UniqueID = "MnuRecAut";
                oMenuParams.String = "Reconciliaciones automaticas";
                oMenuParams.Enabled = true;
                //oMenuParams.Image = sPath + "\\Img\\product1.bmp";
                oMenuParams.Position = countMenu + 1;
                SboMenu.AddEx(oMenuParams);
            }
            catch (Exception ex)
            {
                // CORRECCION: ya no se vuelve a llamar GetFormByTypeAndCount aquí
                // (si la primera llamada falló, ej. SAP cerrándose, repetirla solo
                // vuelve a fallar) y el error ya no se deja escapar del método:
                // solo se informa por la barra de estado.
                Console.WriteLine($"Error Add-On {asm.GetName().Name} : {ex.Message}");
                try { oApp.SetStatusBarMessage("Error Add-On " + asm.GetName().Name + " : " + ex.Message.ToString()); } catch { }
            }
            finally
            {
                // CORRECCION: el des-congelado/actualizacion del formulario del menu
                // ahora va en su propio try/catch. Si SAP ya se está cerrando (u otra
                // falla transitoria), esto ya no tumba el arranque del addon.
                try
                {
                    if (oFormMenu != null)
                    {
                        oFormMenu.Freeze(false);
                        oFormMenu.Update();
                    }
                }
                catch (Exception exFinally)
                {
                    Console.WriteLine($"Error al descongelar formulario de menú (probablemente SAP cerrándose): {exFinally.Message}");
                }

                Ini.ReleaseCOMObjetc(oFormMenu);
                Ini.ReleaseCOMObjetc(SboMenu);
                Ini.ReleaseCOMObjetc(oMenuItem);
                Ini.ReleaseCOMObjetc(oMenuParams);
            }
        }

        /// <summary>
        /// Metodo para agregar filtros a la aplicacion
        /// </summary>
        private void Filters()
        {
            EventFilters oFilters = null;
            EventFilter oFilter = null;
            try
            {


                oFilters = new SAPbouiCOM.EventFilters();

                oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED);
                oFilter.AddEx("143"); // entrada compras
                oFilter.AddEx("721"); // entrada inventario
                oFilter.AddEx("179"); // entrada Nota credito clientes
                oFilter.AddEx("140"); // Salida entrega ventas
                oFilter.AddEx("720"); // Salida Inventarios
                oFilter.AddEx("940"); // Salida Traslados
                oFilter.AddEx("182"); // Salida Devolucion de compras
                oFilter.AddEx("181"); // Salida Nota credito de compras
                oFilter.AddEx("180"); // Entrada Devolucion ventas
                //oFilter.AddEx("133"); // salida factura de ventas
                oFilter.AddEx("141"); // entrada factura de compras
                oFilter.AddEx("65213"); // Salida emision produccion
                oFilter.AddEx("65214"); // entrada recibo de produccion
                oFilter.AddEx("1474000006"); // entrada y salida conteo de inventario
                oFilter.AddEx("1474000001"); // entrada y salida conteo de inventario


                oFilter.AddEx("FRMLTE");
                oFilter.AddEx("FRMLTS");
                oFilter.AddEx("FRMLCON");



                oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK);
                oFilter.AddEx("FRMLTE");
                oFilter.AddEx("FRMLTS");
                oFilter.AddEx("FRMLCON");


                oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD);
                oFilter.AddEx("143"); // entrada compras
                oFilter.AddEx("721"); // entrada inventario
                oFilter.AddEx("179"); // entrada Nota credito clientes
                oFilter.AddEx("140"); // Salida entrega ventas
                oFilter.AddEx("720"); // Salida Inventarios
                oFilter.AddEx("940"); // Salida Traslados
                oFilter.AddEx("182"); //salida Devolucion de compras
                oFilter.AddEx("181"); //salida Nota credito de compras
                oFilter.AddEx("180"); //entrada Devolucion ventas
                //oFilter.AddEx("133"); // salida factura de ventas
                oFilter.AddEx("141"); // entrada factura de compras
                oFilter.AddEx("65213"); // Salida emision produccion
                oFilter.AddEx("65214"); // entrada recibo de produccion
                oFilter.AddEx("1474000006"); // entrada y salida conteo de inventario
                oFilter.AddEx("1474000001"); // entrada y salida conteo de inventario


                oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_CLOSE);
                oFilter.AddEx("143"); // entrada compras
                oFilter.AddEx("721"); // entrada inventario
                oFilter.AddEx("179"); // entrada Nota credito clientes
                oFilter.AddEx("140"); // Salida entrega ventas
                oFilter.AddEx("720"); // Salida Inventarios
                oFilter.AddEx("940"); // Salida Traslados
                oFilter.AddEx("182"); //salida Devolucion de compras
                oFilter.AddEx("181"); //salida Nota credito de compras
                oFilter.AddEx("180"); //entrada Devolucion ventas
                //oFilter.AddEx("133"); // salida factura de ventas
                oFilter.AddEx("141"); // entrada factura de compras
                oFilter.AddEx("65213"); // Salida emision produccion
                oFilter.AddEx("65214"); // entrada recibo de produccion
                oFilter.AddEx("1474000006"); // entrada y salida conteo de inventario
                oFilter.AddEx("1474000001"); // entrada y salida conteo de inventario


                oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD);
                oFilter.AddEx("143"); // entrada compras
                oFilter.AddEx("721"); // entrada inventario
                oFilter.AddEx("179"); // entrada Nota credito clientes
                oFilter.AddEx("140"); // Salida entrega ventas
                oFilter.AddEx("720"); // Salida Inventarios
                oFilter.AddEx("940"); // Salida Traslados
                oFilter.AddEx("182"); //salida Devolucion de compras
                oFilter.AddEx("181"); //salida Nota credito de compras
                oFilter.AddEx("180"); //entrada Devolucion ventas
                //oFilter.AddEx("133"); // salida factura de ventas
                oFilter.AddEx("141"); // entrada factura de compras
                oFilter.AddEx("65213"); // Salida emision produccion
                oFilter.AddEx("65214"); // entrada recibo de produccion
                oFilter.AddEx("1474000006"); // entrada y salida conteo de inventario
                oFilter.AddEx("1474000001"); // entrada y salida conteo de inventario


                oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK);
                oFilter.AddEx("143"); // entrada compras
                oFilter.AddEx("721"); // entrada inventario
                oFilter.AddEx("179"); // entrada Nota credito clientes
                oFilter.AddEx("140"); // Salida entrega ventas
                oFilter.AddEx("720"); // Salida Inventarios
                oFilter.AddEx("940"); // Salida Traslados
                oFilter.AddEx("182"); //salida Devolucion de compras
                oFilter.AddEx("181"); //salida Nota credito de compras
                oFilter.AddEx("180"); //entrada Devolucion ventas
                //oFilter.AddEx("133"); // salida factura de ventas
                oFilter.AddEx("141"); // entrada factura de compras
                oFilter.AddEx("65213"); // Salida emision produccion
                oFilter.AddEx("65214"); // entrada recibo de produccion
                oFilter.AddEx("1474000006"); // entrada y salida conteo de inventario
                oFilter.AddEx("1474000001"); // entrada y salida conteo de inventario

                oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);
                oFilter.AddEx("143"); // entrada compras
                oFilter.AddEx("721"); // entrada inventario
                oFilter.AddEx("179"); // entrada Nota credito clientes
                oFilter.AddEx("140"); // Salida entrega ventas
                oFilter.AddEx("720"); // Salida Inventarios
                oFilter.AddEx("940"); // Salida Traslados
                oFilter.AddEx("182"); //salida Devolucion de compras
                oFilter.AddEx("181"); //salida Nota credito de compras
                oFilter.AddEx("180"); //entrada Devolucion ventas
                //oFilter.AddEx("133"); // salida factura de ventas
                oFilter.AddEx("141"); // entrada factura de compras
                oFilter.AddEx("65213"); // Salida emision produccion
                oFilter.AddEx("65214"); // entrada recibo de produccion
                oFilter.AddEx("1474000006"); // entrada y salida conteo de inventario
                oFilter.AddEx("1474000001"); // entrada y salida conteo de inventario

                oApp.SetFilter(oFilters);

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error Filters Add-On: {asm?.GetName().Name}" + "\n" + ex.Message.ToString() + " " + ex.StackTrace.ToString());
            }
            finally
            {
                if (oFilters != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oFilters);
                }
                if (oFilter != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oFilter);
                }

            }
        }

        /// <summary>
        /// Metodo que inica los eventos en la aplicacion
        /// </summary>
        private void Events()
        {
            oApp.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
            oApp.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            oApp.FormDataEvent += new _IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);
            oApp.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClick);
            oApp.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
            //oApp.StatusBarEvent += new _IApplicationEvents_StatusBarEventEventHandler(SBO_Application_StatusBarEvent);
        }


        #endregion

        #region Eventos globales            
        
        /// <summary>
        ///Metodo de captura de eventos del menu 
        /// </summary>
        /// <param name="pVal">Variable  de evento</param>
        /// <param name="BubbleEvent">Variable de retorno</param>
        private void SBO_Application_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {

            BubbleEvent = true;
            SAPbouiCOM.Form oForm = oApp.Forms.ActiveForm;


            if (!pVal.BeforeAction)
            {
                switch (pVal.MenuUID)
                {
                    case "MnVerLote":
                        List<string> formulariosPermitidos = new List<string>
                                {
                                    "143", // entrada compras
                                    "721", // entrada inventario
                                    "179", // entrada Nota credito clientes
                                    "140", // Salida entrega ventas
                                    "720", // Salida Inventarios
                                    "940", // Salida Traslados
                                    "182", //salida Devolucion de compras
                                    "181", //salida Nota credito de compras
                                    "180", //entrada Devolucion ventas
                                    "141", // entrada factura de compras
                                    "65213", // Salida emision produccion
                                    "65214", // entrada recibo de produccion
                                    "1474000006", // entrada y salida conteo de inventario
                                    "1474000001" // entrada y salida conteo de inventario
                                };
                        if (formulariosPermitidos.Contains(oForm.TypeEx))
                        {
                            switch (oForm.TypeEx)
                            {
                                case "143":
                                    if (formLotesEntrada == null)
                                    {
                                        formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                    }
                                    Tipo = "OPDN";
                                    formLotesEntrada.MostarItems(oForm, Tipo);
                                    break;
                                case "721":
                                    if (formLotesEntrada == null)
                                    {
                                        formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                    }
                                    Tipo = "OIGN";
                                    formLotesEntrada.MostarItems(oForm, Tipo);
                                    break;
                                case "179":
                                    if (formLotesEntrada == null)
                                    {
                                        formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                    }
                                    Tipo = "ORIN";
                                    formLotesEntrada.MostarItems(oForm, Tipo);
                                    break;
                                case "180":
                                    if (formLotesEntrada == null)
                                    {
                                        formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                    }
                                    Tipo = "ORDN";
                                    formLotesEntrada.MostarItems(oForm, Tipo);
                                    break;
                                case "141":
                                    if (formLotesEntrada == null)
                                    {
                                        formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                    }
                                    Tipo = "OPCH";
                                    formLotesEntrada.MostarItems(oForm, Tipo);
                                    break;
                                case "65214":
                                    if (formLotesEntrada == null)
                                    {
                                        formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                    }
                                    Tipo = "OIGN";
                                    formLotesEntrada.MostarItems(oForm, Tipo);
                                    break;
                                case "140":
                                    if (formLotesSalida == null)
                                    {
                                        formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                    }
                                    Tipo = "ODLN";
                                    formLotesSalida.MostarItems(oForm, Tipo);
                                    break;
                                case "720":
                                    if (formLotesSalida == null)
                                    {
                                        formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                    }
                                    Tipo = "OIGE";
                                    formLotesSalida.MostarItems(oForm, Tipo);
                                    break;
                                case "940":
                                    if (formLotesSalida == null)
                                    {
                                        formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                    }
                                    Tipo = "OWTR";
                                    formLotesSalida.MostarItems(oForm, Tipo);
                                    break;
                                case "182":
                                    if (formLotesSalida == null)
                                    {
                                        formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                    }
                                    Tipo = "ORPD";
                                    formLotesSalida.MostarItems(oForm, Tipo);
                                    break;
                                case "181":
                                    if (formLotesSalida == null)
                                    {
                                        formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                    }
                                    Tipo = "ORPC";
                                    formLotesSalida.MostarItems(oForm, Tipo);
                                    break;
                                //case "133":
                                //    if (formLotesSalida == null)
                                //    {
                                //        formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                //    }
                                //    Tipo = "OINV";
                                //    formLotesSalida.MostarItems(oForm, Tipo);
                                //    break;
                                case "65213":
                                    if (formLotesSalida == null)
                                    {
                                        formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                    }
                                    Tipo = "OIGE";
                                    formLotesSalida.MostarItems(oForm, Tipo);
                                    break;
                            }
                        }
                        break;
                }
            }

            Ini.ReleaseCOMObjetc(oForm);


        }


        /// <summary>
        /// Metodo de captura de evnetos de clic derecho
        /// </summary>
        /// <param name="eventInfo">Variable con informacion de evento</param>
        /// <param name="BubbleEvent">Variable de retorno</param>
        private void SBO_Application_RightClick(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oForm = oApp.Forms.Item(eventInfo.FormUID);

            List<string> formulariosPermitidos = new List<string>
            {
               "143", // entrada compras
                "721", // entrada inventario
                "179", // entrada Nota credito clientes
                "140", // Salida entrega ventas
                "720", // Salida Inventarios
                "940", // Salida Traslados
                "182", //salida Devolucion de compras
                "181", //salida Nota credito de compras
                "180", //entrada Devolucion ventas
                "141", // entrada factura de compras
                "65213", // Salida emision produccion
                "65214", // entrada recibo de produccion
                "1474000006", // entrada y salida conteo de inventario
                "1474000001" // entrada y salida conteo de inventario
            };


            if (formulariosPermitidos.Contains(oForm.TypeEx))
            {
                if (eventInfo.BeforeAction.Equals(true))
                {
                    switch (oForm.TypeEx)
                    {
                        case "721":
                            if (eventInfo.ItemUID.Equals("13"))
                            {
                                if (eventInfo.ColUID.Equals("1"))
                                {
                                    addMenuRigthClick();
                                }
                            }
                            break;
                        case "720":
                            if (eventInfo.ItemUID.Equals("13"))
                            {
                                if (eventInfo.ColUID.Equals("1"))
                                {
                                    addMenuRigthClick();
                                }
                            }
                            break;
                        default:
                            if (eventInfo.ItemUID.Equals("38"))
                            {
                                if (eventInfo.ColUID.Equals("1"))
                                {
                                    addMenuRigthClick();
                                }
                            }
                            break;
                    }
                    
                }
                else
                {
                    switch (oForm.TypeEx)
                    {
                        case "721":
                            if (eventInfo.ItemUID.Equals("13"))
                            {
                                deleteMenuRigthClick();
                            }
                            break;
                        case "720":
                            if (eventInfo.ItemUID.Equals("13"))
                            {
                                deleteMenuRigthClick();
                            }
                            break;
                        default:
                            if (eventInfo.ItemUID.Equals("38"))
                            {
                                deleteMenuRigthClick();
                            }
                            break;
                    }
                   
                }
            }

            Ini.ReleaseCOMObjetc(oForm);

        }


        /// <summary>
        /// Eventos para acciones del item
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        /// <exception cref="NotImplementedException"></exception>
        private void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            Console.WriteLine("*******************************");
            Console.WriteLine(pVal.FormUID.ToString());
            Console.WriteLine(pVal.FormType.ToString());
            Console.WriteLine(pVal.FormTypeEx.ToString());
            Console.WriteLine(pVal.EventType.ToString());
            Console.WriteLine(pVal.FormMode.ToString());
            Console.WriteLine(pVal.BeforeAction);


            if (!pVal.BeforeAction)
            {
                switch (pVal.FormTypeEx)
                {
                    case "FRMLTE": //Formulario de entrada de lotes
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_CLICK:
                                switch (pVal.ItemUID)
                                {
                                    case "GrdLineas":
                                        if (pVal.ColUID.Equals("Sel") && !pVal.Row.Equals(-1))
                                        {
                                            if (formLotesEntrada == null)
                                            {
                                                formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                            }
                                            formLotesEntrada.AddInfoLotes(pVal, Tipo);
                                        }

                                        break;
                                    case "BtnAddLn":
                                        if (formLotesEntrada == null)
                                        {
                                            formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                        }
                                        formLotesEntrada.AddLine(pVal);
                                        break;
                                }
                                break;
                        }
                        break;
                    case "FRMLTS": //Formulario de salida de lotes
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_CLICK:
                                switch (pVal.ItemUID)
                                {
                                    case "GrdLineas":
                                        if (pVal.ColUID.Equals("Sel") && !pVal.Row.Equals(-1))
                                        {
                                            if (formLotesSalida == null)
                                            {
                                                formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                            }
                                            formLotesSalida.AddInfoLotes(pVal);
                                        }

                                        break;
                                        //case "GrdDisp":
                                        //    if (pVal.ColUID.Equals("Sel") && !pVal.Row.Equals(-1))
                                        //    {
                                        //        if (formLotesSalida == null)
                                        //        {
                                        //            formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                        //        }
                                        //        //formLotesSalida.AddInfoLotesSalida(pVal);
                                        //    }
                                        //    break;
                                }
                                break;
                        }
                        break;
                    case "143":
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_FORM_CLOSE:
                                if (formLotesEntrada == null)
                                {
                                    formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                }
                                formLotesEntrada.formClose();
                                break;
                                //case BoEventTypes.et_FORM_LOAD:
                                //    if (formLotesEntrada == null)
                                //    {
                                //        formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                //    }
                                //    Tipo = "OPDN";
                                //    formLotesEntrada.LoadControl(pVal, Tipo);
                                //    break;
                        }
                        break;
                    case "140":
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_FORM_CLOSE:
                                if (formLotesSalida == null)
                                {
                                    formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                }
                                formLotesSalida.formClose();
                                break;
                        }
                        break;
                    case "721":
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_FORM_CLOSE:
                                if (formLotesEntrada == null)
                                {
                                    formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                }
                                formLotesEntrada.formClose();
                                break;
                        }
                        break;
                    case "179": // entrada nota credito ventas
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_FORM_CLOSE:
                                if (formLotesEntrada == null)
                                {
                                    formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                }
                                formLotesEntrada.formClose();
                                break;
                        }
                        break;
                    case "182":
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_FORM_CLOSE:
                                if (formLotesSalida == null)
                                {
                                    formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                }
                                formLotesSalida.formClose();
                                break;
                        }
                        break;
                    case "720":
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_FORM_CLOSE:
                                if (formLotesSalida == null)
                                {
                                    formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                }
                                formLotesSalida.formClose();
                                break;
                        }
                        break;
                    case "940":
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_FORM_CLOSE:
                                if (formLotesSalida == null)
                                {
                                    formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                }
                                formLotesSalida.formClose();
                                break;
                        }
                        break;
                    case "181":
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_FORM_CLOSE:
                                if (formLotesSalida == null)
                                {
                                    formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                }
                                formLotesSalida.formClose();
                                break;
                        }
                        break;
                    case "180": //Entrada devolucion ventas
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_FORM_CLOSE:
                                if (formLotesEntrada == null)
                                {
                                    formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                }
                                formLotesEntrada.formClose();
                                break;
                        }
                        break;
                    //case "133": // salida factura de venta
                    //    switch (pVal.EventType)
                    //    {
                    //        case BoEventTypes.et_FORM_CLOSE:
                    //            if (formLotesSalida == null)
                    //            {
                    //                formLotesSalida = new FormLotesSalida(oCompany, oApp);
                    //            }
                    //            formLotesSalida.formClose();
                    //            break;
                    //    }
                    //    break;
                    case "141": //entrada factura compras
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_FORM_CLOSE:
                                if (formLotesEntrada == null)
                                {
                                    formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                }
                                formLotesEntrada.formClose();
                                break;
                        }
                        break;
                    case "65213": //Salida emision de produccion
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_FORM_CLOSE:
                                if (formLotesSalida == null)
                                {
                                    formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                }
                                formLotesSalida.formClose();
                                break;
                        }
                        break;
                    case "65214": // ENTRADA recibo de produccion
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_FORM_CLOSE:
                                if (formLotesEntrada == null)
                                {
                                    formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                }
                                formLotesEntrada.formClose();
                                break;
                        }
                        break;
                    case "1474000001":// ENTRADA recuento de inventario
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_FORM_CLOSE:
                                if (formLotesEntrada == null)
                                {
                                    formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                }
                                formLotesEntrada.formClose();
                                break;
                        }
                        break;
                    case "FRMLCON": //Formulario de entrada de lotes
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_CLICK:
                                switch (pVal.ItemUID)
                                {
                                    case "GrdLineas":
                                        if (pVal.ColUID.Equals("Sel") && !pVal.Row.Equals(-1))
                                        {
                                            if (formLotesConteo == null)
                                            {
                                                formLotesConteo = new FormLotesConteo(oCompany, oApp);
                                            }
                                            formLotesConteo.AddInfoLotes(pVal, Tipo);
                                        }

                                        break;
                                        //case "BtnAddLn":
                                        //    if (formLotesEntrada == null)
                                        //    {
                                        //        formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                        //    }
                                        //    formLotesEntrada.AddLine(pVal);
                                        //    break;
                                }
                                break;
                        }
                        break;
                }
            }
            else
            {
                switch (pVal.FormTypeEx)
                {
                    case "143": //Entrada de mercancía compras.
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_ITEM_PRESSED:
                                switch (pVal.ItemUID)
                                {
                                    case "1":
                                        if (formLotesEntrada == null)
                                        {
                                            formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                        }
                                        Tipo = "OPDN";
                                        BubbleEvent = formLotesEntrada.MostarItems(pVal, Tipo);
                                        break;

                                }

                                break;
                        }
                        break;
                    case "721": //Entrada de mercancía INVENTARIO.
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_ITEM_PRESSED:
                                switch (pVal.ItemUID)
                                {
                                    case "1":
                                        if (formLotesEntrada == null)
                                        {
                                            formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                        }
                                        Tipo = "OIGN";
                                        BubbleEvent = formLotesEntrada.MostarItems(pVal, Tipo);
                                        break;

                                }

                                break;
                        }
                        break;
                    case "179": //entrada Nota credito clientes
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_ITEM_PRESSED:
                                switch (pVal.ItemUID)
                                {
                                    case "1":
                                        if (formLotesEntrada == null)
                                        {
                                            formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                        }
                                        Tipo = "ORIN";
                                        BubbleEvent = formLotesEntrada.MostarItems(pVal, Tipo);
                                        break;

                                }

                                break;
                        }
                        break;
                    case "182": //salida devolucion por compras
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_ITEM_PRESSED:
                                switch (pVal.ItemUID)
                                {
                                    case "1":
                                        if (formLotesSalida == null)
                                        {
                                            formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                        }
                                        Tipo = "ORPD";
                                        BubbleEvent = formLotesSalida.MostarItems(pVal, Tipo);
                                        break;

                                }

                                break;
                        }
                        break;
                    case "FRMLTE":
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_CLICK:
                                switch (pVal.ItemUID)
                                {
                                    case "GrdLineas":
                                        if (pVal.ColUID.Equals("Sel") && !pVal.Row.Equals(-1))
                                        {
                                            if (formLotesEntrada == null)
                                            {
                                                formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                            }
                                            formLotesEntrada.CleanData(pVal);
                                        }

                                        break;
                                }
                                break;
                            case BoEventTypes.et_ITEM_PRESSED:
                                switch (pVal.ItemUID)
                                {
                                    case "1":

                                        if (formLotesEntrada == null)
                                        {
                                            formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                        }
                                        BubbleEvent = formLotesEntrada.AddLotesList(pVal);
                                        break;
                                    case "BtDelR":
                                        if (formLotesEntrada == null)
                                        {
                                            formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                        }
                                        formLotesEntrada.DeleteRow(pVal);
                                        break;
                                }
                                break;
                        }
                        break;
                    case "140": //Entrega de ventas
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_ITEM_PRESSED:
                                switch (pVal.ItemUID)
                                {
                                    case "1":
                                        if (formLotesSalida == null)
                                        {
                                            formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                        }
                                        Tipo = "ODLN";
                                        BubbleEvent = formLotesSalida.MostarItems(pVal, Tipo);
                                        break;
                                        //case "2":
                                        //    if (formLotesSalida == null)
                                        //    {
                                        //        formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                        //    }
                                        //    formLotesSalida.RollBack(pVal, "ODLN");
                                        //    break;
                                }
                                break;
                        }
                        break;
                    case "720": //Salida Inventarios
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_ITEM_PRESSED:
                                switch (pVal.ItemUID)
                                {
                                    case "1":
                                        if (formLotesSalida == null)
                                        {
                                            formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                        }
                                        Tipo = "OIGE";
                                        BubbleEvent = formLotesSalida.MostarItems(pVal, Tipo);
                                        break;
                                        //case "2":
                                        //    if (formLotesSalida == null)
                                        //    {
                                        //        formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                        //    }
                                        //    formLotesSalida.RollBack(pVal, "ODLN");
                                        //    break;
                                }
                                break;
                        }
                        break;
                    case "940": //Salida Traslados
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_ITEM_PRESSED:
                                switch (pVal.ItemUID)
                                {
                                    case "1":
                                        if (formLotesSalida == null)
                                        {
                                            formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                        }
                                        Tipo = "OWTR";
                                        BubbleEvent = formLotesSalida.MostarItems(pVal, Tipo);
                                        break;
                                        //case "2":
                                        //    if (formLotesSalida == null)
                                        //    {
                                        //        formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                        //    }
                                        //    formLotesSalida.RollBack(pVal, "ODLN");
                                        //    break;
                                }
                                break;
                        }
                        break;
                    case "181": //Salida nota credito compras
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_ITEM_PRESSED:
                                switch (pVal.ItemUID)
                                {
                                    case "1":
                                        if (formLotesSalida == null)
                                        {
                                            formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                        }
                                        Tipo = "ORPC";
                                        BubbleEvent = formLotesSalida.MostarItems(pVal, Tipo);
                                        break;
                                        //case "2":
                                        //    if (formLotesSalida == null)
                                        //    {
                                        //        formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                        //    }
                                        //    formLotesSalida.RollBack(pVal, "ODLN");
                                        //    break;
                                }
                                break;
                        }
                        break;
                    case "FRMLTS":
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_CLICK:
                                switch (pVal.ItemUID)
                                {
                                    case "GrdLineas":
                                        if (pVal.ColUID.Equals("Sel") && !pVal.Row.Equals(-1))
                                        {
                                            if (formLotesSalida == null)
                                            {
                                                formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                            }
                                            formLotesSalida.CleanData(pVal);
                                        }

                                        break;
                                }
                                break;
                            case BoEventTypes.et_ITEM_PRESSED:
                                switch (pVal.ItemUID)
                                {
                                    case "BtnAutSel":

                                        if (formLotesSalida == null)
                                        {
                                            formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                        }
                                        formLotesSalida.SalidaLotesAuto(pVal);
                                        break;
                                    case "BtnMan":
                                        if (formLotesSalida == null)
                                        {
                                            formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                        }
                                        formLotesSalida.EnableSelectionManual(pVal);
                                        break;
                                    case "BtnAgre":
                                        if (formLotesSalida == null)
                                        {
                                            formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                        }
                                        formLotesSalida.AddLoteSel(pVal);
                                        break;
                                    case "BtnElim":
                                        if (formLotesSalida == null)
                                        {
                                            formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                        }
                                        formLotesSalida.DelLoteSel(pVal);
                                        break;
                                    case "1":
                                        if (formLotesSalida == null)
                                        {
                                            formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                        }                                       
                                        BubbleEvent =  formLotesSalida.AddLoteSelToList(pVal);
                                        break;
                                }
                                break;

                        }
                        break;
                    case "180": //Entrada devolucion de ventas
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_ITEM_PRESSED:
                                switch (pVal.ItemUID)
                                {
                                    case "1":
                                        if (formLotesEntrada == null)
                                        {
                                            formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                        }
                                        Tipo = "ORDN";
                                        BubbleEvent = formLotesEntrada.MostarItems(pVal, Tipo);
                                        break;

                                }

                                break;
                        }
                        break;
                    //case "133": // salida factura de ventas
                    //    switch (pVal.EventType)
                    //    {
                    //        case BoEventTypes.et_ITEM_PRESSED:
                    //            switch (pVal.ItemUID)
                    //            {
                    //                case "1":
                    //                    if (formLotesSalida == null)
                    //                    {
                    //                        formLotesSalida = new FormLotesSalida(oCompany, oApp);
                    //                    }
                    //                    Tipo = "OINV";
                    //                    BubbleEvent = formLotesSalida.MostarItems(pVal, Tipo);
                    //                    break;
                    //                    //case "2":
                    //                    //    if (formLotesSalida == null)
                    //                    //    {
                    //                    //        formLotesSalida = new FormLotesSalida(oCompany, oApp);
                    //                    //    }
                    //                    //    formLotesSalida.RollBack(pVal, "ODLN");
                    //                    //    break;
                    //            }
                    //            break;
                    //    }
                    //    break;
                    case "141": //Entrada de factura compras.
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_ITEM_PRESSED:
                                switch (pVal.ItemUID)
                                {
                                    case "1":
                                        if (formLotesEntrada == null)
                                        {
                                            formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                        }
                                        Tipo = "OPCH";
                                        BubbleEvent = formLotesEntrada.MostarItems(pVal, Tipo);
                                        break;

                                }

                                break;
                        }
                        break;
                    case "65213": //Salida emision de produccion
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_ITEM_PRESSED:
                                switch (pVal.ItemUID)
                                {
                                    case "1":
                                        if (formLotesSalida == null)
                                        {
                                            formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                        }
                                        Tipo = "OIGE";
                                        BubbleEvent = formLotesSalida.MostarItems(pVal, Tipo);
                                        break;
                                        //case "2":
                                        //    if (formLotesSalida == null)
                                        //    {
                                        //        formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                        //    }
                                        //    formLotesSalida.RollBack(pVal, "ODLN");
                                        //    break;
                                }
                                break;
                        }
                        break;
                    case "65214": //Entrada de recibo de produccion.
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_ITEM_PRESSED:
                                switch (pVal.ItemUID)
                                {
                                    case "1":
                                        if (formLotesEntrada == null)
                                        {
                                            formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                        }
                                        Tipo = "OIGN";
                                        BubbleEvent = formLotesEntrada.MostarItems(pVal, Tipo);
                                        break;

                                }

                                break;
                        }
                        break;
                    case "1474000001": //Recuento de inventario
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_ITEM_PRESSED:
                                switch (pVal.ItemUID)
                                {
                                    case "1470000001":
                                        if (formLotesConteo == null)
                                        {
                                            formLotesConteo = new FormLotesConteo(oCompany, oApp);

                                        }
                                        Tipo = "OINC";
                                        BubbleEvent = formLotesConteo.MostarItems(pVal, Tipo);
                                        break;

                                }

                                break;
                        }
                        break;
                    case "FRMLCON":
                        switch (pVal.EventType)
                        {
                            case BoEventTypes.et_CLICK:
                                switch (pVal.ItemUID)
                                {
                                    case "GrdLineas":
                                        if (pVal.ColUID.Equals("Sel") && !pVal.Row.Equals(-1))
                                        {
                                            if (formLotesConteo == null)
                                            {
                                                formLotesConteo = new FormLotesConteo(oCompany, oApp);
                                            }
                                            formLotesConteo.CleanData(pVal);
                                        }

                                        break;
                                }
                                break;
                            case BoEventTypes.et_ITEM_PRESSED:
                                switch (pVal.ItemUID)
                                {
                                    case "1":

                                        if (formLotesConteo == null)
                                        {
                                            formLotesConteo = new FormLotesConteo(oCompany, oApp);
                                        }
                                        BubbleEvent = formLotesConteo.AddLotesList(pVal);
                                        break;
                                    case "BtnAgre":
                                        if (formLotesConteo == null)
                                        {
                                            formLotesConteo = new FormLotesConteo(oCompany, oApp);
                                        }
                                        formLotesConteo.AddLoteSel(pVal);
                                        break;
                                    case "BtnElim":
                                        if (formLotesConteo == null)
                                        {
                                            formLotesConteo = new FormLotesConteo(oCompany, oApp);
                                        }
                                        formLotesConteo.DelLoteSel(pVal);
                                        break;
                                }
                                break;
                        }
                        break;
                }

            }
        }


        /// <summary>
        /// Eventos para acciones de la aplicación
        /// </summary>
        /// <param name="EventType"></param>
        /// <exception cref="NotImplementedException"></exception>
        private void SBO_Application_AppEvent(BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case BoAppEventTypes.aet_ShutDown:
                    System.Windows.Forms.Application.Exit();
                    break;
                case BoAppEventTypes.aet_ServerTerminition:
                    System.Windows.Forms.Application.Exit();
                    break;
                case BoAppEventTypes.aet_CompanyChanged:
                    System.Windows.Forms.Application.Exit();
                    break;
                case BoAppEventTypes.aet_FontChanged:
                    System.Windows.Forms.Application.Exit();
                    break;
                case BoAppEventTypes.aet_LanguageChanged:
                    System.Windows.Forms.Application.Exit();
                    break;
            }
        }


        /// <summary>
        /// Eventos de Status bar
        /// </summary>
        /// <param name="Text"></param>
        /// <param name="messageType"></param>
        //private void SBO_Application_StatusBarEvent(string Text, BoStatusBarMessageType messageType)
        //{
        //    String s = Text;
        //    Text = "";
        //    messageType = BoStatusBarMessageType.smt_None;

        //}


        /// <summary>
        /// Eventos de datos de formulario
        /// </summary>
        /// <param name="BusinessObjectInfo"></param>
        /// <param name="BubbleEvent"></param>
        private void SBO_Application_FormDataEvent(ref BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            Console.WriteLine("*******************************");
            Console.WriteLine(BusinessObjectInfo.FormUID.ToString());
            Console.WriteLine(BusinessObjectInfo.FormTypeEx.ToString());
            Console.WriteLine(BusinessObjectInfo.EventType.ToString());
            Console.WriteLine(BusinessObjectInfo.BeforeAction);
            Console.WriteLine(BusinessObjectInfo.ActionSuccess);


            if (!BusinessObjectInfo.BeforeAction && BusinessObjectInfo.ActionSuccess)
            {
                // CORRECCION (aprobaciones): cuando el documento entra a un proceso de
                // aprobación, lo que SAP realmente agrega primero es un Borrador
                // (Draft, Type = "112"), NO el documento definitivo. "esBorrador" marca
                // ese caso: en vez de crear el lote como CONFIRMADO, se crea como
                // PENDIENTE. Cuando el borrador se aprueba y se agrega como documento
                // real (puede ser en otra sesión/computador, por ejemplo la del
                // aprobador), llega un segundo et_FORM_DATA_ADD con el Type real:
                // ese caso ya NO es esBorrador, y ahí se confirma el lote pendiente
                // (por Tipo de documento + Numero de documento) o, si no hay nada
                // pendiente (flujo normal sin aprobación), se crea directo como
                // CONFIRMADO. Toda esa lógica vive en ProcesarAdd() de cada formulario.
                bool esBorrador =
                    BusinessObjectInfo.EventType.Equals(BoEventTypes.et_FORM_DATA_ADD)
                    && BusinessObjectInfo.Type.Equals(((int)SAPbobsCOM.BoObjectTypes.oDrafts).ToString());

                switch (BusinessObjectInfo.FormTypeEx)
                {
                    case "143": //Entrada de mercancía compras.
                        switch (BusinessObjectInfo.EventType)
                        {
                            case BoEventTypes.et_FORM_DATA_ADD:
                                if (formLotesEntrada == null)
                                {
                                    formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                }
                                BubbleEvent = formLotesEntrada.ProcesarAdd(esBorrador, "OPDN");

                                break;
                            case BoEventTypes.et_FORM_DATA_LOAD:
                                if (formLotesEntrada == null)
                                {
                                    formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                }
                                formLotesEntrada.formClose();

                                break;
                        }
                        break;
                    case "721": //Entrada de mercancía INVENTARIO.
                        switch (BusinessObjectInfo.EventType)
                        {
                            case BoEventTypes.et_FORM_DATA_ADD:
                                if (formLotesEntrada == null)
                                {
                                    formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                }
                                BubbleEvent = formLotesEntrada.ProcesarAdd(esBorrador, "OIGN");

                                break;
                        }
                        break;
                    case "179": //entrada Nota credito clientes
                        switch (BusinessObjectInfo.EventType)
                        {
                            case BoEventTypes.et_FORM_DATA_ADD:
                                if (formLotesEntrada == null)
                                {
                                    formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                }
                                BubbleEvent = formLotesEntrada.ProcesarAdd(esBorrador, "ORIN");

                                break;
                            case BoEventTypes.et_FORM_DATA_LOAD:
                                if (formLotesEntrada == null)
                                {
                                    formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                }
                                formLotesEntrada.formClose();

                                break;
                        }
                        break;
                    case "182": //salida devolucion por compras
                        switch (BusinessObjectInfo.EventType)
                        {
                            case BoEventTypes.et_FORM_DATA_ADD:
                                if (formLotesSalida == null)
                                {
                                    formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                }
                                BubbleEvent = formLotesSalida.ProcesarAdd(esBorrador, "ORPD");

                                break;
                            case BoEventTypes.et_FORM_DATA_LOAD:
                                if (formLotesSalida == null)
                                {
                                    formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                }
                                formLotesSalida.formClose();

                                break;
                        }
                        break;
                    case "140": //Entrega de ventas.
                        switch (BusinessObjectInfo.EventType)
                        {
                            case BoEventTypes.et_FORM_DATA_ADD:
                                if (formLotesSalida == null)
                                {
                                    formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                }
                                BubbleEvent = formLotesSalida.ProcesarAdd(esBorrador, "ODLN");

                                break;
                            case BoEventTypes.et_FORM_DATA_LOAD:
                                if (formLotesSalida == null)
                                {
                                    formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                }
                                formLotesSalida.formClose();

                                break;
                        }
                        break;
                    case "720": //Salida Inventarios
                        switch (BusinessObjectInfo.EventType)
                        {
                            case BoEventTypes.et_FORM_DATA_ADD:
                                if (formLotesSalida == null)
                                {
                                    formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                }
                                BubbleEvent = formLotesSalida.ProcesarAdd(esBorrador, "OIGE");

                                break;
                        }
                        break;
                    case "940": //Salida Traslados
                        switch (BusinessObjectInfo.EventType)
                        {
                            case BoEventTypes.et_FORM_DATA_ADD:
                                if (formLotesSalida == null)
                                {
                                    formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                }
                                BubbleEvent = formLotesSalida.ProcesarAdd(esBorrador, "OWTR");

                                break;
                            case BoEventTypes.et_FORM_DATA_LOAD:
                                if (formLotesSalida == null)
                                {
                                    formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                }
                                formLotesSalida.formClose();

                                break;
                        }
                        break;
                    case "181": //Salida nota credito compras
                        switch (BusinessObjectInfo.EventType)
                        {
                            case BoEventTypes.et_FORM_DATA_ADD:
                                if (formLotesSalida == null)
                                {
                                    formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                }
                                BubbleEvent = formLotesSalida.ProcesarAdd(esBorrador, "ORPC");

                                break;
                            case BoEventTypes.et_FORM_DATA_LOAD:
                                if (formLotesSalida == null)
                                {
                                    formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                }
                                formLotesSalida.formClose();

                                break;
                        }
                        break;
                    case "180": //Entrada devolucion ventas
                        switch (BusinessObjectInfo.EventType)
                        {
                            case BoEventTypes.et_FORM_DATA_ADD:
                                if (formLotesEntrada == null)
                                {
                                    formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                }
                                BubbleEvent = formLotesEntrada.ProcesarAdd(esBorrador, "ORDN");

                                break;
                            case BoEventTypes.et_FORM_DATA_LOAD:
                                if (formLotesEntrada == null)
                                {
                                    formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                }
                                formLotesEntrada.formClose();

                                break;
                        }
                        break;
                    //case "133": //salida factura de ventas.
                    //    switch (BusinessObjectInfo.EventType)
                    //    {
                    //        case BoEventTypes.et_FORM_DATA_ADD:
                    //            if (formLotesSalida == null)
                    //            {
                    //                formLotesSalida = new FormLotesSalida(oCompany, oApp);
                    //            }
                    //            BubbleEvent = formLotesSalida.SalidaLotes();

                    //            break;
                    //        case BoEventTypes.et_FORM_DATA_LOAD:
                    //            if (formLotesSalida == null)
                    //            {
                    //                formLotesSalida = new FormLotesSalida(oCompany, oApp);
                    //            }
                    //            formLotesSalida.formClose();

                    //            break;
                    //    }
                    //    break;
                    case "141": //Entrada de factura compras.
                        switch (BusinessObjectInfo.EventType)
                        {
                            case BoEventTypes.et_FORM_DATA_ADD:
                                if (formLotesEntrada == null)
                                {
                                    formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                }
                                BubbleEvent = formLotesEntrada.ProcesarAdd(esBorrador, "OPCH");

                                break;
                            case BoEventTypes.et_FORM_DATA_LOAD:
                                if (formLotesEntrada == null)
                                {
                                    formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                }
                                formLotesEntrada.formClose();

                                break;
                        }
                        break;
                    case "65213": //salida emision de produccion
                        switch (BusinessObjectInfo.EventType)
                        {
                            case BoEventTypes.et_FORM_DATA_ADD:
                                if (formLotesSalida == null)
                                {
                                    formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                }
                                BubbleEvent = formLotesSalida.ProcesarAdd(esBorrador, "OIGE");

                                break;
                            case BoEventTypes.et_FORM_DATA_LOAD:
                                if (formLotesSalida == null)
                                {
                                    formLotesSalida = new FormLotesSalida(oCompany, oApp);
                                }
                                formLotesSalida.formClose();

                                break;
                        }
                        break;
                    case "65214": //Entrada de recibo de produccion.
                        switch (BusinessObjectInfo.EventType)
                        {
                            case BoEventTypes.et_FORM_DATA_ADD:
                                if (formLotesEntrada == null)
                                {
                                    formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                }
                                BubbleEvent = formLotesEntrada.ProcesarAdd(esBorrador, "OIGN");

                                break;
                            case BoEventTypes.et_FORM_DATA_LOAD:
                                if (formLotesEntrada == null)
                                {
                                    formLotesEntrada = new FormLotesEntrada(oCompany, oApp);
                                }
                                formLotesEntrada.formClose();

                                break;
                        }
                        break;
                    case "1474000001": //Entrada de reconteo de inventario.
                        switch (BusinessObjectInfo.EventType)
                        {
                            case BoEventTypes.et_FORM_DATA_ADD:
                                if (formLotesConteo == null)
                                {
                                    formLotesConteo = new FormLotesConteo(oCompany, oApp);
                                }
                                BubbleEvent = formLotesConteo.CrearLotes();

                                break;
                            case BoEventTypes.et_FORM_DATA_LOAD:
                                if (formLotesConteo == null)
                                {
                                    formLotesConteo = new FormLotesConteo(oCompany, oApp);
                                }
                                formLotesConteo.formClose();

                                break;
                        }
                        break;
                }
            }
            else
            {


            }
        }

        #endregion

        #region Metodos globales

        /// <summary>
        /// Libera objetos COM en tiempo de ejecución.
        /// </summary>
        /// <param name="obj">Objeto COM a liberar.</param>
        public static void ReleaseCOMObjetc(object obj)
        {
            try
            {
                if (obj != null && System.Runtime.InteropServices.Marshal.IsComObject(obj))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            }
            catch (Exception ex)
            {
                // CORRECCION: ya no se muestra MessageBox (causaba el mensaje "fantasma"
                // en documentos donde el formulario de lotes no estaba abierto) y ya no se
                // fuerza GC.Collect()/WaitForPendingFinalizers() en cada liberación
                // (causaba las desconexiones esporádicas del add-on obligatorio).
                Console.WriteLine($"Error en liberación de objeto COM: {ex?.Message}");
            }
        }

        public static double ConvertToSystemFormat(string sapValue)
        {
            var settings = oCompany.GetCompanyService().GetAdminInfo();
            string sapDecimalSep = settings.DecimalSeparator;
            string sapThousandsSep = settings.ThousandsSeparator;
            string sysDecimalSep = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;

            // Eliminar separadores de miles y reemplazar decimal SAP por el del sistema
            string normalizedValue = sapValue
                .Replace(sapThousandsSep, string.Empty)
                .Replace(sapDecimalSep, sysDecimalSep);



            if (!double.TryParse(normalizedValue, NumberStyles.Any, CultureInfo.CurrentCulture.NumberFormat, out double number))
                throw new Exception($"Error al convertir el valor '{sapValue}' a número con formato del sistema.");

            return number;
        }

        public static string ConvertToSapFormat(double number, int decimals = 2)
        {
            var settings = oCompany.GetCompanyService().GetAdminInfo();
            string sapDecimalSep = settings.DecimalSeparator;
            string sapThousandsSep = settings.ThousandsSeparator;

            NumberFormatInfo sapFormat = new NumberFormatInfo
            {
                NumberDecimalSeparator = sapDecimalSep,
                NumberGroupSeparator = sapThousandsSep
            };

            return number.ToString($"N{decimals}", sapFormat);
        }
        #endregion


        #region "Metodos privados"
        /// <summary>
        /// Metodo para crear el menu de clic derecho
        /// </summary>
        private void addMenuRigthClick()
        {
            SAPbouiCOM.MenuItem oMenuItem = null;
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuCreationParams oMenuCreationParams = null;
            try
            {
                oMenuCreationParams = (SAPbouiCOM.MenuCreationParams)oApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oMenuCreationParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oMenuCreationParams.UniqueID = "MnVerLote";
                oMenuCreationParams.String = "Ver lotes";
                oMenuCreationParams.Enabled = true;
                oMenuItem = oApp.Menus.Item("1280");
                oMenus = oMenuItem.SubMenus;
                if (oMenus.Exists(oMenuCreationParams.UniqueID))
                    return;
                oMenus.AddEx(oMenuCreationParams);
            }
            catch (Exception ex)
            {
                oApp.MessageBox("Error: Add-On Lotes hornitos" + " " + ex.Message.ToString() + " " + ex.StackTrace.ToString());
            }
            finally
            {
                ReleaseCOMObjetc(oMenuItem);
                ReleaseCOMObjetc(oMenus);
                ReleaseCOMObjetc(oMenuCreationParams);
            }
        }

        private void deleteMenuRigthClick()
        {
            try
            {
                SAPbouiCOM.MenuItem menuItem = oApp.Menus.Item("1280");
                SAPbouiCOM.Menus menus = menuItem.SubMenus;
                if (menus.Exists("MnVerLote"))
                    menus.RemoveEx("MnVerLote");
            }
            catch (Exception ex)
            {
                oApp.MessageBox("Error: " + ex.Message.ToString() + " " + ex.StackTrace.ToString());
            }
        }
        #endregion
    }
}
