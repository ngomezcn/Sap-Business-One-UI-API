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
        private SAPbouiCOM.Form oForm;

        public void ConnectToSBO()
        {
            try
            {
                SAPbouiCOM.SboGuiApi sboGuiApi = null;
                sboGuiApi = new SAPbouiCOM.SboGuiApi();

                sboGuiApi.Connect("0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056");

                SBO_Application = sboGuiApi.GetApplication(-1);

                SBO_Application.StatusBar.SetText("Add-on cargado correctamente", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se ha podido conectar con sap");
            }
        }

        public SBO_UI_API()
        {
            ConnectToSBO();
            //bindEventHandlersToSBO();

            CrearFormularioSimple();
            oForm.Visible = true;
            GuardarComoXML();
        }

        private void bindEventHandlersToSBO()
        {
            // Eventos gestionados por SBO_Application, se debe pasar el evento tal como se especifica en la documentación
            // Para ver todos los tipos de eventos que gestiona SBO ver la interfaz: SAPbouiCOM._IApplicationEvents_Event 

            SBO_Application.AppEvent  += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
            SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
            SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
        }
        
        // Gestiona los eventos a nivel de aplicacion SBO.
        // +info sobre los eventos de aplicacion: Help Center -> [BoAppEventTypes Enumeration] 
        private void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType) 
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

        // Gestiona los eventos de interacción con los menús -> [Barra superior, Barra de herramientas y Menú general]
        // Aclaracion:
        // 1. Un evento de menu se considera una interacción con uno de los botones para realizar una gestión, ej: abrir formulario inter. comerc., asistente de generar PDF, abrir calendario...
        // 2. No se considera evento de menú, ej: desplegar un submenú en la barra superior o abrir una carpeta en el menú general...
        // 3. Tampoco se considera evento de menú, interactuar con la posible ui de gestión abierto desde el menú, ej: formulario inter. comerc., asistente de generar PDF, calendario...
        private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent) 
        {
            BubbleEvent = true; 
            if (pVal.BeforeAction == true) // Podemos gestionar el evento antes que SAP
            {
                SBO_Application.SetStatusBarMessage("Menu item: " + pVal.MenuUID + " enviado ANTES de que SAP lo procese.", SAPbouiCOM.BoMessageTime.bmt_Long, true);

                // Este código se ejecutaria antes de que SBO procese el evento
                // ...
                // ... 

                BubbleEvent = true; // Podemos evitar que SBO llegue a procesar el evento si ponemos BubbleEvent = false
                // Extra: SAP propone que a través de esta vía podemos mostrar nuestros propios formularios en lugar de los predefinidos por SBO
            }
            else
            {
                SBO_Application.SetStatusBarMessage("Menu item: " + pVal.MenuUID + " enviado DESPUES de que SAP lo procese.", SAPbouiCOM.BoMessageTime.bmt_Long, true);

                // Este código se ejecutaría después de que SBO procese el evento
                // ...
                // ... 
            }
        }

        // Todos los eventos son en esencia un item event (a excepcion de unos muy concretos)
        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if(pVal.FormType != 0)
            {
                SAPbouiCOM.BoEventTypes EventEnum = pVal.EventType; // Help Center -> [BoEventTypes Enumeration] 

                // Aquí se podrían indicar validaciones para filtrar los eventos que nos interesan: ej: Al dejar de editar el field Nombre del form inter. comerc.
                // ...
                // ...

                Console.Write("Un evento tipo " + EventEnum.ToString() + " ha sido enviado desde un formulario con ID: " + FormUID + "\n");
            }
        }

        private void CrearFormularioSimple()
        {
            SAPbouiCOM.Item oItem = null;

            SAPbouiCOM.Button oButton = null;
            SAPbouiCOM.StaticText oStaticText = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.ComboBox oComboBox = null;

            SAPbouiCOM.FormCreationParams oCreationParams = null;

            // Creando parametros del formulario
            oCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.UniqueID = "dfga";

            // Creando formulario a partir de los parametros indicados
            
            oForm = SBO_Application.Forms.AddEx(oCreationParams);

            // Creando los datos de entrada (esto todavia no muestra nada hay que bindearlo a un item)
            oForm.DataSources.UserDataSources.Add("EditSource", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("CombSource", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);

            // Algunas propiedades del nuevo formulario
            oForm.Title = "Simple Form";
            oForm.Left = 400;
            oForm.Top = 100;
            oForm.ClientHeight = 80;
            oForm.ClientWidth = 350;

            // SBO gestionara automaticamente los eventos si ponemos el UID a 1
            // ** Añadimos boton de OK **
            oItem = oForm.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 6;
            oItem.Width = 65;
            oItem.Top = 51;
            oItem.Height = 19;
            oButton = ((SAPbouiCOM.Button)(oItem.Specific));
            oButton.Caption = "Ok";

            // SBO gestionara automaticamente los eventos si ponemos el UID a 2
            // ** Añadimos boton de Cancelar ** 
            oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 75;
            oItem.Width = 65;
            oItem.Top = 51;
            oItem.Height = 19;
            oButton = ((SAPbouiCOM.Button)(oItem.Specific));
            oButton.Caption = "Cancel";

            // ** Añadimos un rectangulo **
            oItem = oForm.Items.Add("Rect1", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oItem.Left = 0;
            oItem.Width = 344;
            oItem.Top = 1;
            oItem.Height = 49;

            // ** Añadimos Texto Estatico **
            oItem = oForm.Items.Add("StaticTxt1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = 7;
            oItem.Width = 148;
            oItem.Top = 8;
            oItem.Height = 14;
            oItem.LinkTo = "EditText1";

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Static Text 1";

            // ** Añadimos otro Texto Estatico **
            oItem = oForm.Items.Add("StaticTxt2", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = 7;
            oItem.Width = 148;
            oItem.Top = 24;
            oItem.Height = 14;
            oItem.LinkTo = "ComboBox1";

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStaticText.Caption = "Static Text 2";

            // ** Añadimos Edit Text **
            oItem = oForm.Items.Add("EditText1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = 157;
            oItem.Width = 163;
            oItem.Top = 8;
            oItem.Height = 14;

            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
            oEditText.DataBind.SetBound(true, "", "EditSource"); // Bindeamos el Edit Text a los datos de entrada creados anteriormente 

            oEditText.String = "Edit Text 1";

            // ** Añadimos Combo Box **
            oItem = oForm.Items.Add("ComboBox1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oItem.Left = 157;
            oItem.Width = 163;
            oItem.Top = 24;
            oItem.Height = 14;
            oItem.DisplayDesc = false;

            oComboBox = ((SAPbouiCOM.ComboBox)(oItem.Specific));
            oComboBox.DataBind.SetBound(true, "", "CombSource"); // Bindeamos el Edit Text a los datos de entrada creados anteriormente 

            oComboBox.ValidValues.Add("1", "Combo Value 1");
            oComboBox.ValidValues.Add("2", "Combo Value 2");
            oComboBox.ValidValues.Add("3", "Combo Value 3");
        }

        private void GuardarComoXML()
        {
            System.Xml.XmlDocument oXmlDoc = null;

            oXmlDoc = new System.Xml.XmlDocument();

            string sXmlString = null;

            // get the form as an XML string
            sXmlString = oForm.GetAsXML();

            // load the form's XML string to the
            // XML document object
            oXmlDoc.LoadXml(sXmlString);

            string sPath = null;

            sPath = System.IO.Directory.GetParent(Application.StartupPath).ToString();

            oXmlDoc.Save((sPath + @"\MySimpleForm.xml"));
        }
    }
}
