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
            bindEventHandlersToSBO();
        }

        public void bindEventHandlersToSBO()
        {
            // Eventos gestionados por SBO_Application, se debe pasar el evento tal como se especifica en la documentación
            // Para ver todos los tipos de eventos gestiona SBO ver la interfaz: SAPbouiCOM._IApplicationEvents_Event 

            SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
            SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
        }
        
        // Recibe los eventos a nivel de aplicacion SBO. Ver -> [BoAppEventTypes Enumeration] 
        public void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType) 
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown: 
                    SBO_Application.MessageBox("El evento de cerrar SBO se ha detectado" + Environment.NewLine + "Cerrando Add On", 1, "Listo");

                    // Aqui debe ir el codigo para gestionar el apagado del Addon antes de que se cierre SAP
                    // ...
                    // ...

                    System.Windows.Forms.Application.Exit();
                    break;

                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged: // 
                    SBO_Application.MessageBox("Un evento de cambio empresa ha sido detectado");

                    if (SBO_Application.Company.Name == "nombreEmpresa") // Podemos identificar la empresa a la que se esta cambiando
                    {
                         // ...
                         // ...
                    }
                    break;

                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    SBO_Application.MessageBox("Se ha detectado una modificacion del lenguaje");
                    
                    // ...
                    // ...

                    break;
            }
        }

        // Recibe los eventos de interacción con los menús -> [Barra superior, Barra de herramientas y Menú general]
        // Aclaracion:
        // 1. Un evento de menu se considera una interacción con uno de los botones para realizar una gestión, ej: abrir formulario inter. comerc., asistente de generar PDF, abrir calendario...
        // 2. No se considera evento de menú, ej: desplegar un submenú en la barra superior o abrir una carpeta en el menú general...
        // 3. Tampoco se considera evento de menú, interactuar con la posible ui de gestión abierto desde el menú, ej: formulario inter. comerc., asistente de generar PDF, calendario...
        private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent) 
        {
            BubbleEvent = false;
            if (pVal.BeforeAction == true) // Podemos gestionar el evento antes que SAP
            {
                SBO_Application.SetStatusBarMessage("Menu item: " + pVal.MenuUID + " enviado ANTES de que SAP lo procese.", SAPbouiCOM.BoMessageTime.bmt_Long, true);

                // Este código se ejecutaria antes de que SBO procese el evento
                // ...
                // ... 

                BubbleEvent = true; // Podemos evitar que SBO llegue a procesar el evento si ponemos BubbleEvent = false
            }
            else
            {
                SBO_Application.SetStatusBarMessage("Menu item: " + pVal.MenuUID + " enviado DESPUES de que SAP lo procese.", SAPbouiCOM.BoMessageTime.bmt_Long, true);

                //  Este código se ejecutaria despues de que SBO procese el evento
                // ...
                // ... 
            }
        }
    }
}
