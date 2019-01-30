
using System;
using System.Collections.Generic;

using MYOB.AccountRight.SDK;
using MYOB.AccountRight.SDK.Contracts;
using System.Configuration;
using System.IO;

namespace MyOBCustomService.Helpers
{
    public static class SessionManager
    {

        private const string companyFile = "MyCompanyFile";
        private const string iApiConfiguration = "MyConfiguration";
        private const string iCompanyFileCredentials = "MyCredentials";
        private const string iOAuthKeyService = "MyOAuthKeyService";
        private const string companyFieles = "CompanyFiles";
        private const string selectedCompanyFile = "CompanyFile";



        public static string CsOAuthServer = "https://secure.myob.com/oauth2/account/authorize/";

        public static string CsOAuthScope = "CompanyFile";
        public static string LocalApiUrl = "http://localhost:8080/accountright";
        public static string DeveloperKey = ConfigurationManager.AppSettings["DeveloperKey"];
        public static string DeveloperSecret = ConfigurationManager.AppSettings["DeveloperSecret"];
        public static string Code = ConfigurationManager.AppSettings["Code"];
        public static string CompanyId = ConfigurationManager.AppSettings["CompanyId"];
        public static string CompanyUserId = ConfigurationManager.AppSettings["CompanyUserId"];
        public static string CompanyPassword = ConfigurationManager.AppSettings["CompanyPassword"];
        public static DateTime ToDate =DateTime.Now;
        public static int TimerInHours =Convert.ToInt16( ConfigurationManager.AppSettings["TimerInHours"]);
        public static string FilePath = ConfigurationManager.AppSettings["FilePath"];



        public static CompanyFile CompanyFile
        { get;set;
            //get
            //{
            //    if (HttpContext.Current.Session[companyFile] != null)
            //    {
            //        return (CompanyFile)HttpContext.Current.Session[companyFile];
            //    }
            //    else
            //    {
            //        return null;
            //    }
            //}
            //set
            //{
            //    HttpContext.Current.Session[companyFile] = value;
            //}
        }
        public static IApiConfiguration MyConfiguration
        {
            get; set;
            //get
            //{
            //    if (HttpContext.Current.Session[iApiConfiguration] != null)
            //    {
            //        return (IApiConfiguration)HttpContext.Current.Session[iApiConfiguration];
            //    }
            //    else
            //    {
            //        return null;
            //    }
            //}
            //set
            //{
            //    HttpContext.Current.Session[iApiConfiguration] = value;
            //}
        }

        public static ICompanyFileCredentials MyCredentials
        {
            get; set;
            //get
            //{
            //    if (HttpContext.Current.Session[iCompanyFileCredentials] != null)
            //    {
            //        return (ICompanyFileCredentials)HttpContext.Current.Session[iCompanyFileCredentials];
            //    }
            //    else
            //    {
            //        return null;
            //    }
            //}
            //set
            //{
            //    HttpContext.Current.Session[iCompanyFileCredentials] = value;
            //}
        }
        public static IOAuthKeyService MyOAuthKeyService
        {
            get; set;
            //get
            //{
            //    if (HttpContext.Current.Session[iOAuthKeyService] != null)
            //    {
            //        return (IOAuthKeyService)HttpContext.Current.Session[iOAuthKeyService];
            //    }
            //    else
            //    {
            //        return null;
            //    }
            //}
            //set
            //{
            //    HttpContext.Current.Session[iOAuthKeyService] = value;
            //}
        }

        public static List<CompanyFile> CompanyFiles
        {
            get; set;
            //get
            //{
            //    if (HttpContext.Current.Session[companyFieles] == null)
            //    {
            //        HttpContext.Current.Session[companyFieles] = new List<CompanyFile>();
            //    }
            //    return (List<CompanyFile>)HttpContext.Current.Session[companyFieles];
            //}
            //set { HttpContext.Current.Session[companyFieles] = value; }

        }

        public static CompanyFile SelectedCompanyFile
        {
            get; set;
            //get
            //{
            //    if (HttpContext.Current.Session[selectedCompanyFile] == null)
            //    {
            //        HttpContext.Current.Session[selectedCompanyFile] = new CompanyFile();
            //    }
            //    return (CompanyFile)HttpContext.Current.Session[selectedCompanyFile];
            //}
            //set { HttpContext.Current.Session[selectedCompanyFile] = value; }

        }
    }
}
