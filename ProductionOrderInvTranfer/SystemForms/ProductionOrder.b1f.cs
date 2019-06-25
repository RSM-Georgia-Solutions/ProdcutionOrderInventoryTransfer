using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM.Framework;
using BoFormObjectEnum = SAPbouiCOM.BoFormObjectEnum;
using BoMessageTime = SAPbouiCOM.BoMessageTime;
using Form = SAPbouiCOM.Form;

namespace ProductionOrderInvTranfer
{
    [FormAttribute("65211", "SystemForms/ProductionOrder.b1f")]
    class ProductionOrder : SystemFormBase
    {
        public ProductionOrder()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_0").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Button Button0;

        private void OnCustomInitialize()
        {

        }

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Form orderForm = Application.SBO_Application.Forms.ActiveForm;
            string docEntry = orderForm.DataSources.DBDataSources.Item("OWOR").GetValue("DocEntry", 0);
            if (string.IsNullOrWhiteSpace(docEntry))
            {
                Application.SBO_Application.SetStatusBarMessage("დოკუმენტი არ არის დამატებული",
                    BoMessageTime.bmt_Short, true);
                return;
            }

            ProductionOrders productionOrder = (ProductionOrders)DiManager.Company.GetBusinessObject(BoObjectTypes.oProductionOrders);
            productionOrder.GetByKey(int.Parse(docEntry));

            StockTransfer inventoryTransfer = (StockTransfer)DiManager.Company.GetBusinessObject(BoObjectTypes.oStockTransfer);

            try
            {
                var transferDocEntry = productionOrder.UserFields.Fields.Item("U_TransferDocEntry").Value.ToString();
                if (!string.IsNullOrWhiteSpace(transferDocEntry) && transferDocEntry != "0")
                {
                    Application.SBO_Application.SetStatusBarMessage("უკვე გამოწერილია",
                        BoMessageTime.bmt_Short, true);
                    return;
                }
            }
            catch (Exception)
            {
                Application.SBO_Application.SetStatusBarMessage("TransferDocEntry არ არის დამატებული",
                    BoMessageTime.bmt_Short, true);
                return;
            }
            string wareHouse = string.Empty;
            try
            {
                wareHouse = productionOrder.UserFields.Fields.Item("U_WareHouse").Value.ToString();
            }
            catch (Exception)
            {
                Application.SBO_Application.SetStatusBarMessage("WareHouse არ არის დამატებული",
                    BoMessageTime.bmt_Short, true);
                return;
            }




            inventoryTransfer.DocDate = productionOrder.PostingDate;
            inventoryTransfer.TaxDate = productionOrder.PostingDate;


            for (int i = 0; i < productionOrder.Lines.Count; i++)
            {
                double qunatity;
                productionOrder.Lines.SetCurrentLine(i);
                try
                {
                    qunatity = double.Parse(productionOrder.Lines.UserFields.Fields.Item("U_TransferQuantity").Value.ToString(), CultureInfo.InvariantCulture);
                }
                catch (Exception)
                {
                    Application.SBO_Application.SetStatusBarMessage("TransferQuantity არ არის დამატებული",
                        BoMessageTime.bmt_Short, true);
                    return;
                }

                string itemCode = productionOrder.Lines.ItemNo;

                inventoryTransfer.Lines.ItemCode = itemCode;
                inventoryTransfer.Lines.Quantity = qunatity;
                inventoryTransfer.FromWarehouse = productionOrder.Lines.Warehouse;
                inventoryTransfer.ToWarehouse = wareHouse;
                inventoryTransfer.Lines.Add();
            }

            var res = inventoryTransfer.Add();
            if (res != 0)
            {
                Application.SBO_Application.SetStatusBarMessage(DiManager.Company.GetLastErrorDescription(),
                    BoMessageTime.bmt_Short, true);
            }
            else
            {
                productionOrder.UserFields.Fields.Item("U_TransferDocEntry").Value =
                    DiManager.Company.GetNewObjectKey();
                productionOrder.Update();
                Application.SBO_Application.ActivateMenuItem("1304");
            }
        }
    }
}
