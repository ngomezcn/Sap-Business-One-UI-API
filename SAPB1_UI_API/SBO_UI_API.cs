using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SAPB1_UI_API
{
    internal class SBO_UI_API
    {
        private SAPbouiCOM.Application SBO_Application;

        public void ConnectToSBO()
        {
            try
            {
                SAPbouiCOM.SboGuiApi sboGuiApi = null;
                sboGuiApi = new SAPbouiCOM.SboGuiApi();

                sboGuiApi.Connect("0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056");

                SBO_Application = sboGuiApi.GetApplication(-1);

                SBO_Application.StatusBar.SetText("Add-on cargado correctamente", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                SBO_Application.MessageBox("Add-on cargado correctametne");

                

            }
            catch (Exception ex)
            {
                MessageBox.Show("No se ha podido conectar con sap");
            }
        }

        public SBO_UI_API()
        {
            ConnectToSBO();

            SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);

            SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
        }

        public void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    SBO_Application.MessageBox("El evento de cerrar SBO se ha detectado" + Environment.NewLine + "Cerrando Add On", 1, "Listo");

                    // Aqui debe ir el codigo para gestionar el apagado del Addon

                    System.Windows.Forms.Application.Exit();
                    break;

                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    SBO_Application.MessageBox("Un evento de empresa ha sido detectado");
                    Console.Write("Company change");

                    if (SBO_Application.Company.Name == "nombreEmpresa")   
                    {
                         // ... 
                         // ...
                         // ...
                    }
                    break;

                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    SBO_Application.MessageBox("Se ha detectado una modificacion del lenguaje");
                    break;
            }
        }

        private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = false;
            if (pVal.BeforeAction == true) // Podemos escoger si gestionar el evento antes que SAP
            {
                SBO_Application.SetStatusBarMessage("Menu item: " + pVal.MenuUID + " sent an event BEFORE SAP Business One processes it.", SAPbouiCOM.BoMessageTime.bmt_Long, true);

                // Este codigo se ejecutara antes de que SBO haya procesado el evento
                // ...
                // ... 

                BubbleEvent = false; // Ademas podemos evitar que SBO procese el evento si ponemos BubbleEvent = false
            }
            else
            {
                SBO_Application.SetStatusBarMessage("Menu item: " + pVal.MenuUID + " sent an event AFTER SAP Business One processes it.", SAPbouiCOM.BoMessageTime.bmt_Long, true);

                // Este codigo se ejecutara despues de que SBO haya procesado el evento
                // ...
                // ... 
            }
        }
    }
}
