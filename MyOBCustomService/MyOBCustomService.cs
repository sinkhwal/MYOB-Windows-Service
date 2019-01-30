using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using MYOB.AccountRight.SDK;
using MyOBCustomService.Helpers;
using MYOB.AccountRight.SDK.Contracts;
using MYOB.Model;

namespace MyOBCustomService
{
    public partial class MyOBCustomService : ServiceBase
    {
        private int PageSize = 1000;
        private int eventId = 1;
        private static IApiConfiguration _configurationCloud;
        private static IOAuthKeyService _oAuthKeyService;
        public MyOBCustomService()
        {
            InitializeComponent();
            eventLog1 = new System.Diagnostics.EventLog();
            System.Diagnostics.EventLog.DeleteEventSource("MyOBCustomService");
            if (!System.Diagnostics.EventLog.SourceExists("MyOBCustomService"))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                    "MyOBCustomService", "MyOBCustomServiceLog");
            }
            eventLog1.Source = "MyOBCustomService";
            eventLog1.Log = "MyOBCustomServiceLog";
        }

        protected override void OnStart(string[] args)
        {
           // Update the service state to Start Pending.
           ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            serviceStatus.dwWaitHint = 300000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 60000*60*SessionManager.TimerInHours; // 60 seconds
            timer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnTimer);
            timer.Start();
            _configurationCloud = new ApiConfiguration(SessionManager.DeveloperKey, SessionManager.DeveloperSecret, "http://localhost:60669/Default");
            _oAuthKeyService = new OAuthKeyService();
            //if (_oAuthKeyService.OAuthResponse == null)
            //{
                var oauthService = new MYOB.AccountRight.SDK.Services.OAuthService(_configurationCloud);
                _oAuthKeyService.OAuthResponse = oauthService.GetTokens(SessionManager.Code);
          //  }

            SessionManager.MyOAuthKeyService = _oAuthKeyService;
            SessionManager.MyConfiguration = _configurationCloud;

            var cfsCloud = new MYOB.AccountRight.SDK.Services.CompanyFileService(SessionManager.MyConfiguration, null, SessionManager.MyOAuthKeyService);
            CompanyFile[] companyFiles = cfsCloud.GetRange();
            CompanyFile companyFile = companyFiles.FirstOrDefault(a => a.Id == Guid.Parse(SessionManager.CompanyId));
            SessionManager.SelectedCompanyFile = companyFile;
            ICompanyFileCredentials credentials = new CompanyFileCredentials(SessionManager.CompanyUserId, SessionManager.CompanyPassword);
            SessionManager.MyCredentials = credentials;
            // OnTimer();
            // Update the service state to Running.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
            OnTimer(null,null);
        }

        protected override void OnStop()
        {
            eventLog1.WriteEntry("In OnStop.");
        }
        protected override void OnContinue()
        {
            eventLog1.WriteEntry("In OnContinue.");
        }

        public void OnTimer(object sender, System.Timers.ElapsedEventArgs args)
       // public void OnTimer()
        {
            // TODO: Insert monitoring activities here.
            eventLog1.WriteEntry("Monitoring the System", EventLogEntryType.Information, eventId++);

           
            HandleSales();
            HandlePurchase();


        }


        protected void HandleSales()
        {
            
            string filter = string.Format("$filter=Date ge datetime'{0}' and Date le datetime'{1}'", SessionManager.ToDate.AddYears(-1).ToString("yyyy-MM-dd"), SessionManager.ToDate.ToString("yyyy-MM-dd"));
            string pageFilter = string.Empty;

              List<SalesData> listPL = new List<SalesData>();
             listPL.Clear();
            var service = new MYOB.AccountRight.SDK.Services.Sale.ItemOrderService(SessionManager.MyConfiguration, null, SessionManager.MyOAuthKeyService);

            int count = 1000;
            for (int currentPage = 1; count >= 1000; currentPage++)
            {
                pageFilter = string.Format("&$top={0}&$skip={1}&$orderby=Date desc", PageSize,
                                         PageSize * (currentPage - 1));

                var list = service.GetRange(SessionManager.SelectedCompanyFile, filter + pageFilter, SessionManager.MyCredentials, null);
                count = list.Items.Count();
                //var invoisvc = new ItemInvoiceService(SessionManager.MyConfiguration, null, SessionManager.MyOAuthKeyService);
                //var list = invoisvc.GetRange(SessionManager.SelectedCompanyFile, null, SessionManager.MyCredentials);
                //
                var invoices = list.Items;
                foreach (var inv in invoices)
                {
                    // var items = inv.Lines;
                    foreach (var item in inv.Lines)
                    {
                        listPL.Add(new SalesData
                        {
                            CustomerName = inv.Customer != null ? inv.Customer.Name : "",
                            TransactionNumber = inv.Number,
                            TransactionDate = inv.Date != null ? inv.Date.ToString("yyyy-MM-dd") : "",
                            TransactionType = "Sales Order",
                            TransactionStatus = inv.Status.ToString(),
                            Itemumber = item.Item != null ? item.Item.Number : "",
                            Product = item.Item != null ? item.Item.Name : "",
                            //  AccountNumber=inv.ac,
                            LineMemo = inv.JournalMemo,
                            EmployeeName = inv.Salesperson != null ? inv.Salesperson.Name : "",
                            Qty = item.ShipQuantity,
                            // TaxExAmount = item.Total,
                            Total = item.Total,
                            TaxCode = item.TaxCode != null ? item.TaxCode.Code.ToString() : "",
                            PromisedDate = inv.PromisedDate,
                            ItemName = item.Description


                        });
                    }
                }
            }

            var serviceItemInvoicService = new MYOB.AccountRight.SDK.Services.Sale.ItemInvoiceService(SessionManager.MyConfiguration, null, SessionManager.MyOAuthKeyService);
            count = 1000;
            for (int currentPage = 1; count >= 1000; currentPage++)
            {
                pageFilter = string.Format("&$top={0}&$skip={1}&$orderby=Date desc", PageSize,
                                         PageSize * (currentPage - 1));

                var list = serviceItemInvoicService.GetRange(SessionManager.SelectedCompanyFile, filter + pageFilter, SessionManager.MyCredentials, null);
                count = list.Items.Count();
                //var invoisvc = new ItemInvoiceService(SessionManager.MyConfiguration, null, SessionManager.MyOAuthKeyService);
                //var list = invoisvc.GetRange(SessionManager.SelectedCompanyFile, null, SessionManager.MyCredentials);
                //
                var invoices = list.Items;
                foreach (var inv in invoices)
                {
                    // var items = inv.Lines;
                    foreach (var item in inv.Lines)
                    {
                        listPL.Add(new SalesData
                        {
                            CustomerName = inv.Customer != null ? inv.Customer.Name : "",
                            TransactionNumber = inv.Number,
                            TransactionDate = inv.Date != null ? inv.Date.ToString("yyyy-MM-dd") : "",
                            TransactionType = "Sales Order",
                            TransactionStatus = inv.Status.ToString(),
                            Itemumber = item.Item != null ? item.Item.Number : "",
                            Product = item.Item != null ? item.Item.Name : "",
                            //  AccountNumber=inv.ac,
                            LineMemo = inv.JournalMemo,
                            EmployeeName = inv.Salesperson != null ? inv.Salesperson.Name : "",
                            Qty = item.ShipQuantity,
                            // TaxExAmount = item.Total,
                            Total = item.Total,
                            TaxCode = item.TaxCode != null ? item.TaxCode.Code.ToString() : "",
                            PromisedDate = inv.PromisedDate,
                            ItemName = item.Description

                        });
                    }
                }
            }

            CreateExcelFile.CreateExcelDocument(listPL,System.IO.Path.Combine(SessionManager.FilePath, "Sales-" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx"));

        }

        protected void HandlePurchase()
        {
   string filter = string.Format("$filter=Date ge datetime'{0}' and Date le datetime'{1}'", SessionManager.ToDate.AddYears(-1).ToString("yyyy-MM-dd"), SessionManager.ToDate.ToString("yyyy-MM-dd"));
            string pageFilter = string.Empty;
            var service = new MYOB.AccountRight.SDK.Services.Purchase.ItemBillService(SessionManager.MyConfiguration, null, SessionManager.MyOAuthKeyService);
            //   var service = new ItemOrderService(SessionManager.MyConfiguration, null, SessionManager.MyOAuthKeyService);
            List<SalesData> listPL = new List<SalesData>();

            listPL.Clear();
            int count = 1000;
            for (int currentPage = 1; count >= 1000; currentPage++)
            {
                pageFilter = string.Format("&$top={0}&$skip={1}&$orderby=Date desc", PageSize,
                                         PageSize * (currentPage - 1));

                var list = service.GetRange(SessionManager.SelectedCompanyFile, filter + pageFilter, SessionManager.MyCredentials, null);
                count = list.Items.Count();
                //var invoisvc = new ItemInvoiceService(SessionManager.MyConfiguration, null, SessionManager.MyOAuthKeyService);
                //var list = invoisvc.GetRange(SessionManager.SelectedCompanyFile, null, SessionManager.MyCredentials);
                //
                var invoices = list.Items;
                foreach (var inv in invoices)
                {
                    var salesData = new SalesData();
                    // var items = inv.Lines;
                    salesData.Total = inv.TotalAmount;
                    salesData.TaxCode = inv.FreightTaxCode.Code.ToString();
                    salesData.PromisedDate = inv.PromisedDate;
                    salesData.CustomerName = inv.Supplier.Name;
                    salesData.TransactionNumber = inv.Number;
                    salesData.TransactionDate = inv.Date.ToString("yyyy-MM-dd");
                    // TransactionType=inv.InvoiceType,
                    salesData.TransactionStatus = inv.Status.ToString();
                    if (inv.Lines != null)
                        foreach (var item in inv.Lines)
                        {


                            salesData.Itemumber = item.Item != null ? item.Item.Number : "";
                            salesData.ItemName = item.Item != null ? item.Item.Name : "";
                            // item.//  AccountNumber=inv.ac,
                            salesData.LineMemo = inv.JournalMemo;
                            //  item./ EmployeeName = inv.Salesperson != null ? inv.Salesperson.Name : "",
                            salesData.Qty = item.BillQuantity;



                        }
                    listPL.Add(salesData);
                }
            }

            var serviceItemInvoicService = new MYOB.AccountRight.SDK.Services.Purchase.ItemPurchaseOrderService(SessionManager.MyConfiguration, null, SessionManager.MyOAuthKeyService);
            count = 1000;
            for (int currentPage = 1; count >= 1000; currentPage++)
            {
                pageFilter = string.Format("&$top={0}&$skip={1}&$orderby=Date desc", PageSize,
                                         PageSize * (currentPage - 1));

                var list = serviceItemInvoicService.GetRange(SessionManager.SelectedCompanyFile, filter + pageFilter, SessionManager.MyCredentials, null);
                count = list.Items.Count();
                //var invoisvc = new ItemInvoiceService(SessionManager.MyConfiguration, null, SessionManager.MyOAuthKeyService);
                //var list = invoisvc.GetRange(SessionManager.SelectedCompanyFile, null, SessionManager.MyCredentials);
                //
                var invoices = list.Items;
                foreach (var inv in invoices)
                {
                    // var items = inv.Lines;
                    foreach (var item in inv.Lines)
                    {
                        var saleDataInvItem = new SalesData();
                        // var items = inv.Lines;
                        saleDataInvItem.Total = inv.TotalAmount;
                        saleDataInvItem.TaxCode = inv.FreightTaxCode.Code.ToString();
                        saleDataInvItem.PromisedDate = inv.PromisedDate;
                        saleDataInvItem.CustomerName = inv.Supplier.Name;
                        saleDataInvItem.TransactionNumber = inv.Number;
                        saleDataInvItem.TransactionDate = inv.Date.ToString("yyyy-MM-dd");
                        // TransactionType=inv.InvoiceType,
                        saleDataInvItem.TransactionStatus = inv.Status.ToString();
                        foreach (var itemInvoic in inv.Lines)
                        {


                            saleDataInvItem.Itemumber = itemInvoic.Item != null ? itemInvoic.Item.Number : "";
                            saleDataInvItem.ItemName = itemInvoic.Item != null ? itemInvoic.Item.Name : "";
                            // item.//  AccountNumber=inv.ac,
                            saleDataInvItem.LineMemo = inv.JournalMemo;
                            //  item./ EmployeeName = inv.Salesperson != null ? inv.Salesperson.Name : "",
                            saleDataInvItem.Qty = itemInvoic.BillQuantity;



                        }
                        listPL.Add(saleDataInvItem);
                    }
                }
            }
            CreateExcelFile.CreateExcelDocument(listPL, System.IO.Path.Combine(SessionManager.FilePath, "Purchases-" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx"));
        }

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(System.IntPtr handle, ref ServiceStatus serviceStatus);

        internal void OnDebug()
        {
            OnStart(null);
        }

        public enum ServiceState
        {
            SERVICE_STOPPED = 0x00000001,
            SERVICE_START_PENDING = 0x00000002,
            SERVICE_STOP_PENDING = 0x00000003,
            SERVICE_RUNNING = 0x00000004,
            SERVICE_CONTINUE_PENDING = 0x00000005,
            SERVICE_PAUSE_PENDING = 0x00000006,
            SERVICE_PAUSED = 0x00000007,
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct ServiceStatus
        {
            public int dwServiceType;
            public ServiceState dwCurrentState;
            public int dwControlsAccepted;
            public int dwWin32ExitCode;
            public int dwServiceSpecificExitCode;
            public int dwCheckPoint;
            public int dwWaitHint;
        };
    }
}
