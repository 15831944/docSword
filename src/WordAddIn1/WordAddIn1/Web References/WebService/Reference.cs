﻿//------------------------------------------------------------------------------
// <auto-generated>
//     此代码由工具生成。
//     运行时版本:4.0.30319.42000
//
//     对此文件的更改可能会导致不正确的行为，并且如果
//     重新生成代码，这些更改将会丢失。
// </auto-generated>
//------------------------------------------------------------------------------

// 
// 此源代码是由 Microsoft.VSDesigner 4.0.30319.42000 版自动生成。
// 
#pragma warning disable 1591

namespace OfficeAssist.WebService {
    using System;
    using System.Web.Services;
    using System.Diagnostics;
    using System.Web.Services.Protocols;
    using System.Xml.Serialization;
    using System.ComponentModel;
    
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1586.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Web.Services.WebServiceBindingAttribute(Name="WordAddinServiceSoap", Namespace="http://tempuri.org/")]
    public partial class WordAddinService : System.Web.Services.Protocols.SoapHttpClientProtocol {
        
        private System.Threading.SendOrPostCallback ValidFileDateOperationCompleted;
        
        private System.Threading.SendOrPostCallback DownFileOperationCompleted;
        
        private System.Threading.SendOrPostCallback GetUrlOperationCompleted;
        
        private System.Threading.SendOrPostCallback GetFilesOperationCompleted;
        
        private System.Threading.SendOrPostCallback ActiveProjectOperationCompleted;
        
        private System.Threading.SendOrPostCallback SignForPersonOperationCompleted;
        
        private System.Threading.SendOrPostCallback SignForEntireOperationCompleted;
        
        private bool useDefaultCredentialsSetExplicitly;
        
        /// <remarks/>
        public WordAddinService() {
            this.Url = global::OfficeAssist.Properties.Settings.Default.docSword_WebService_WordAddinService;
            if ((this.IsLocalFileSystemWebService(this.Url) == true)) {
                this.UseDefaultCredentials = true;
                this.useDefaultCredentialsSetExplicitly = false;
            }
            else {
                this.useDefaultCredentialsSetExplicitly = true;
            }
        }
        
        public new string Url {
            get {
                return base.Url;
            }
            set {
                if ((((this.IsLocalFileSystemWebService(base.Url) == true) 
                            && (this.useDefaultCredentialsSetExplicitly == false)) 
                            && (this.IsLocalFileSystemWebService(value) == false))) {
                    base.UseDefaultCredentials = false;
                }
                base.Url = value;
            }
        }
        
        public new bool UseDefaultCredentials {
            get {
                return base.UseDefaultCredentials;
            }
            set {
                base.UseDefaultCredentials = value;
                this.useDefaultCredentialsSetExplicitly = true;
            }
        }
        
        /// <remarks/>
        public event ValidFileDateCompletedEventHandler ValidFileDateCompleted;
        
        /// <remarks/>
        public event DownFileCompletedEventHandler DownFileCompleted;
        
        /// <remarks/>
        public event GetUrlCompletedEventHandler GetUrlCompleted;
        
        /// <remarks/>
        public event GetFilesCompletedEventHandler GetFilesCompleted;
        
        /// <remarks/>
        public event ActiveProjectCompletedEventHandler ActiveProjectCompleted;
        
        /// <remarks/>
        public event SignForPersonCompletedEventHandler SignForPersonCompleted;
        
        /// <remarks/>
        public event SignForEntireCompletedEventHandler SignForEntireCompleted;
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/ValidFileDate", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string ValidFileDate(string dateString) {
            object[] results = this.Invoke("ValidFileDate", new object[] {
                        dateString});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public void ValidFileDateAsync(string dateString) {
            this.ValidFileDateAsync(dateString, null);
        }
        
        /// <remarks/>
        public void ValidFileDateAsync(string dateString, object userState) {
            if ((this.ValidFileDateOperationCompleted == null)) {
                this.ValidFileDateOperationCompleted = new System.Threading.SendOrPostCallback(this.OnValidFileDateOperationCompleted);
            }
            this.InvokeAsync("ValidFileDate", new object[] {
                        dateString}, this.ValidFileDateOperationCompleted, userState);
        }
        
        private void OnValidFileDateOperationCompleted(object arg) {
            if ((this.ValidFileDateCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.ValidFileDateCompleted(this, new ValidFileDateCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/DownFile", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        [return: System.Xml.Serialization.XmlElementAttribute(DataType="base64Binary")]
        public byte[] DownFile(string fileName, string Key) {
            object[] results = this.Invoke("DownFile", new object[] {
                        fileName,
                        Key});
            return ((byte[])(results[0]));
        }
        
        /// <remarks/>
        public void DownFileAsync(string fileName, string Key) {
            this.DownFileAsync(fileName, Key, null);
        }
        
        /// <remarks/>
        public void DownFileAsync(string fileName, string Key, object userState) {
            if ((this.DownFileOperationCompleted == null)) {
                this.DownFileOperationCompleted = new System.Threading.SendOrPostCallback(this.OnDownFileOperationCompleted);
            }
            this.InvokeAsync("DownFile", new object[] {
                        fileName,
                        Key}, this.DownFileOperationCompleted, userState);
        }
        
        private void OnDownFileOperationCompleted(object arg) {
            if ((this.DownFileCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.DownFileCompleted(this, new DownFileCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetUrl", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string GetUrl(string fileName, string Key) {
            object[] results = this.Invoke("GetUrl", new object[] {
                        fileName,
                        Key});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public void GetUrlAsync(string fileName, string Key) {
            this.GetUrlAsync(fileName, Key, null);
        }
        
        /// <remarks/>
        public void GetUrlAsync(string fileName, string Key, object userState) {
            if ((this.GetUrlOperationCompleted == null)) {
                this.GetUrlOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetUrlOperationCompleted);
            }
            this.InvokeAsync("GetUrl", new object[] {
                        fileName,
                        Key}, this.GetUrlOperationCompleted, userState);
        }
        
        private void OnGetUrlOperationCompleted(object arg) {
            if ((this.GetUrlCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetUrlCompleted(this, new GetUrlCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetFiles", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string GetFiles(string key) {
            object[] results = this.Invoke("GetFiles", new object[] {
                        key});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public void GetFilesAsync(string key) {
            this.GetFilesAsync(key, null);
        }
        
        /// <remarks/>
        public void GetFilesAsync(string key, object userState) {
            if ((this.GetFilesOperationCompleted == null)) {
                this.GetFilesOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetFilesOperationCompleted);
            }
            this.InvokeAsync("GetFiles", new object[] {
                        key}, this.GetFilesOperationCompleted, userState);
        }
        
        private void OnGetFilesOperationCompleted(object arg) {
            if ((this.GetFilesCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetFilesCompleted(this, new GetFilesCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/ActiveProject", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string ActiveProject(string UserName, string ActionCode, string MachineCode) {
            object[] results = this.Invoke("ActiveProject", new object[] {
                        UserName,
                        ActionCode,
                        MachineCode});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public void ActiveProjectAsync(string UserName, string ActionCode, string MachineCode) {
            this.ActiveProjectAsync(UserName, ActionCode, MachineCode, null);
        }
        
        /// <remarks/>
        public void ActiveProjectAsync(string UserName, string ActionCode, string MachineCode, object userState) {
            if ((this.ActiveProjectOperationCompleted == null)) {
                this.ActiveProjectOperationCompleted = new System.Threading.SendOrPostCallback(this.OnActiveProjectOperationCompleted);
            }
            this.InvokeAsync("ActiveProject", new object[] {
                        UserName,
                        ActionCode,
                        MachineCode}, this.ActiveProjectOperationCompleted, userState);
        }
        
        private void OnActiveProjectOperationCompleted(object arg) {
            if ((this.ActiveProjectCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.ActiveProjectCompleted(this, new ActiveProjectCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/SignForPerson", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string SignForPerson(string MachineCode, bool isSign) {
            object[] results = this.Invoke("SignForPerson", new object[] {
                        MachineCode,
                        isSign});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public void SignForPersonAsync(string MachineCode, bool isSign) {
            this.SignForPersonAsync(MachineCode, isSign, null);
        }
        
        /// <remarks/>
        public void SignForPersonAsync(string MachineCode, bool isSign, object userState) {
            if ((this.SignForPersonOperationCompleted == null)) {
                this.SignForPersonOperationCompleted = new System.Threading.SendOrPostCallback(this.OnSignForPersonOperationCompleted);
            }
            this.InvokeAsync("SignForPerson", new object[] {
                        MachineCode,
                        isSign}, this.SignForPersonOperationCompleted, userState);
        }
        
        private void OnSignForPersonOperationCompleted(object arg) {
            if ((this.SignForPersonCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.SignForPersonCompleted(this, new SignForPersonCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/SignForEntire", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string SignForEntire(string MachineCode, bool isSign) {
            object[] results = this.Invoke("SignForEntire", new object[] {
                        MachineCode,
                        isSign});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public void SignForEntireAsync(string MachineCode, bool isSign) {
            this.SignForEntireAsync(MachineCode, isSign, null);
        }
        
        /// <remarks/>
        public void SignForEntireAsync(string MachineCode, bool isSign, object userState) {
            if ((this.SignForEntireOperationCompleted == null)) {
                this.SignForEntireOperationCompleted = new System.Threading.SendOrPostCallback(this.OnSignForEntireOperationCompleted);
            }
            this.InvokeAsync("SignForEntire", new object[] {
                        MachineCode,
                        isSign}, this.SignForEntireOperationCompleted, userState);
        }
        
        private void OnSignForEntireOperationCompleted(object arg) {
            if ((this.SignForEntireCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.SignForEntireCompleted(this, new SignForEntireCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        public new void CancelAsync(object userState) {
            base.CancelAsync(userState);
        }
        
        private bool IsLocalFileSystemWebService(string url) {
            if (((url == null) 
                        || (url == string.Empty))) {
                return false;
            }
            System.Uri wsUri = new System.Uri(url);
            if (((wsUri.Port >= 1024) 
                        && (string.Compare(wsUri.Host, "localHost", System.StringComparison.OrdinalIgnoreCase) == 0))) {
                return true;
            }
            return false;
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1586.0")]
    public delegate void ValidFileDateCompletedEventHandler(object sender, ValidFileDateCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1586.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class ValidFileDateCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal ValidFileDateCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1586.0")]
    public delegate void DownFileCompletedEventHandler(object sender, DownFileCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1586.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class DownFileCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal DownFileCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public byte[] Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((byte[])(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1586.0")]
    public delegate void GetUrlCompletedEventHandler(object sender, GetUrlCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1586.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetUrlCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal GetUrlCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1586.0")]
    public delegate void GetFilesCompletedEventHandler(object sender, GetFilesCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1586.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetFilesCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal GetFilesCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1586.0")]
    public delegate void ActiveProjectCompletedEventHandler(object sender, ActiveProjectCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1586.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class ActiveProjectCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal ActiveProjectCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1586.0")]
    public delegate void SignForPersonCompletedEventHandler(object sender, SignForPersonCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1586.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class SignForPersonCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal SignForPersonCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1586.0")]
    public delegate void SignForEntireCompletedEventHandler(object sender, SignForEntireCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1586.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class SignForEntireCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal SignForEntireCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string)(this.results[0]));
            }
        }
    }
}

#pragma warning restore 1591