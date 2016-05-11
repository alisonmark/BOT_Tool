using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Diagnostics;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using System.Configuration;
using HttpRequestHelper;
using System.Drawing;
using System.Globalization;
using Visa_Appointment_Request;
using CaptchaUtils;
using System.Net;
using System.Threading.Tasks;

namespace Auto_VAR
{
    public partial class VarItem
    {
        #region Fields

        private static int _captchaProviderIndex = 0;
        public static int CaptchaProviderIndex
        {
            get { return _captchaProviderIndex; }
            set { _captchaProviderIndex = value; }
        }

        private static int _countManual = 0;
        public static int CountManual { get; set; }

        public string CaptchaManual { get; set; }

        private DateTime _lastRun = DateTime.MinValue;
        private DateTime _lastCheckedTime = DateTime.MinValue;

        public object Tag { get; set; }

        private object _lockObj = new object();

        private Stopwatch _st = new Stopwatch();
        private TimeSpan _delay = TimeSpan.MinValue;
        private TimeSpan _diffTime = TimeSpan.MinValue;

        private bool _isRunning = false;
        public bool IsRunning
        {
            get { return _isRunning; }
            set { _isRunning = value; }
        }

        private Random _rand = new Random();

        private List<KeyValueItem> _appointDate = new List<KeyValueItem>();
        private VarObject _item;

        public VarObject Item
        {
            get { lock (_lockObj) { return _item; } }
            set { lock (_lockObj) { _item = value; } }
        }

        private FpEventHandler<string> _showMessage;
        public FpEventHandler<string> ShowMessage
        {
            get { if (_showMessage == null) _showMessage = (e, s) => { }; return _showMessage; }
            set { _showMessage = value; }
        }

        private FpEventHandler<bool> _decaptchaFinish;
        public FpEventHandler<bool> DecaptchaFinish
        {
            get { if (_decaptchaFinish == null) _decaptchaFinish = (e, s) => { }; return _decaptchaFinish; }
            set { _decaptchaFinish = value; }
        }

        private FpEventHandler<Bitmap> _startManualCaptcha;
        public FpEventHandler<Bitmap> StartManualCaptcha
        {
            get { if (_startManualCaptcha == null) _startManualCaptcha = (e, s) => { }; return _startManualCaptcha; }
            set { _startManualCaptcha = value; }
        }

        public string CurrentProxy { get; set; }

        #endregion

        public VarItem(VarObject item)
        {
            this._item = item;
        }

        public VarItem() { }

        #region Step 1 -> Step 7

        public void RunBySession()
        {
            bool isResetSession = false;
            if (_lastCheckedTime.AddMinutes(Setting.SessionTime) < DateTime.Now)
            {
                isResetSession = true;
                _lastCheckedTime = DateTime.Now;
            }

            if (!isResetSession
                && DateTime.Now < _lastRun.AddMinutes(Setting.SessionTimeout)
                && (_item.Status == VarState.NotOpen || _item.Status == VarState.InvalidCaptcha))
            {
                PreviousStep2();
            }
            else
            {
                Step1_SelectEmbassy();
            }
        }

        // Step 1: Chọn quốc tịch
        public void Step1_SelectEmbassy()
        {
            _item.Status = VarState.Running;
            if (Setting.UseProxy)
            {
                http.Proxy = GetRandomProxy();
                CurrentProxy = string.Format("{0}:{1}", http.Proxy.Address.Host, http.Proxy.Address.Port);
            }
            else
            {
                http.Proxy = null;
                CurrentProxy = string.Empty;
            }

            _lastRun = DateTime.Now;

            _item.NotOpenTime = string.Empty;
            _item.OpenTime = string.Empty;
            try
            {
                ShowMessage(this, "Step 1: Đang tải https://visapoint.eu/form...");
                string postData = string.Empty;

                _lastResponse = http.FetchHttpGet("https://visapoint.eu/form", string.Empty);

                if (http.LastUrl.Contains("disclaimer"))
                {
                    ShowMessage(this, "Step 1: Đang tải https://visapoint.eu/disclaimer...");
                    postData = string.Format("rsm1_TSM={0}&__EVENTTARGET=ctl00$cp1$btnAccept&__EVENTARGUMENT=&__VIEWSTATE={1}&__VIEWSTATEENCRYPTED=&__EVENTVALIDATION={2}"
                                            + "&ctl00%24ddLocale=English+%28United+Kingdom%29"
                                            + " &ctl00_ddLocale_ClientState="
                                            + "&ctl00_cp1_btnDecline_ClientState="
                                            + "&ctl00_cp1_btnAccept_ClientState="
                                            + "&ctl00_cp1_rbClose_ClientState="
                                            + "&ctl00_cp1_rttDecline_ClientState=", GetRsm1(_lastResponse), GetViewState(_lastResponse), GetEventValidation(_lastResponse));

                    _lastResponse = http.FetchHttpPost("https://visapoint.eu/disclaimer", "https://visapoint.eu/disclaimer", postData);

                    //ShowMessage(this, "Step 1: Đang tải https://visapoint.eu/action...");
                    //_lastResponse = http.FetchHttpGet("https://visapoint.eu/action", "https://visapoint.eu/disclaimer");

                    ShowMessage(this, "Step 1: Đang tải https://visapoint.eu/form...");
                    _lastResponse = http.FetchHttpGet("https://visapoint.eu/form", http.LastUrl);
                }

                //ParseCitizenship(response);
                //ParseResidence(response);

                ShowMessage(this, string.Format("Step 1: Đang chọn quốc tịch: {0}...", _item.Citizenship.Value));
                WriteLog(string.Format("Step 1: Đang chọn quốc tịch: {0}...", _item.Citizenship.Value));
                _lastResponse = RequestCitizenship(_lastResponse);
                //ParseEmbassy(_lastResponse);
                ShowMessage(this, "Step 1: Đã chọn xong quốc tịch:  " + _item.Citizenship.Value);
                WriteLog("Step 1: Đã chọn xong quốc tịch:  " + _item.Citizenship.Value);

                Step2_SelectPurposeOfStay();
            }
            catch (Exception ex)
            {
                _item.Status = VarState.Error;
                _item.ErrorMsg = ex.Message;
                LogUtils.WriteLog("Step 1: Chọn quốc tịch gặp lỗi : " + ex.StackTrace, "err");
                ShowMessage(this, "Step 1: Chọn quốc tịch gặp lỗi : " + ex.Message);
                WriteLog("Step 1: Chọn quốc tịch gặp lỗi : " + ex.Message);
            }
        }

        // Step 2: Chọn Visa
        public void Step2_SelectPurposeOfStay()
        {
            try
            {
                ShowMessage(this, string.Format("Step 2: Đang chọn Mục đích lưu trú: {0}...", _item.PurposeOfStay.Value));
                WriteLog(string.Format("Step 2: Đang chọn Mục đích lưu trú: {0}...", _item.PurposeOfStay.Value));
                //ParseVisaType(_lastResponse);
                _lastResponse = RequestPurposeOfStay(_lastResponse);
                if (http.LastUrl.Contains("disclaimer"))
                {
                    Step1_SelectEmbassy();
                    ShowMessage(this, string.Format("Step 2: Đang chọn Mục đích lưu trú: {0}...", _item.PurposeOfStay.Value));
                    WriteLog(string.Format("Step 2: Đang chọn Mục đích lưu trú: {0}...", _item.PurposeOfStay.Value));
                    _lastResponse = RequestPurposeOfStay(_lastResponse);
                }
                ShowMessage(this, "Step 2: Đã chọn xong Mục đích lưu trú: " + _item.PurposeOfStay.Value);
                WriteLog("Step 2: Đã chọn xong Mục đích lưu trú: " + _item.PurposeOfStay.Value);

                RefreshCaptcha(_lastResponse);

                // Note: Không xử lý captcha luôn, đợi đến thời gian chỉ định thì xử lý
                //Step3_SendCaptcha();
            }
            catch (Exception ex)
            {
                _item.Status = VarState.Error;
                _item.ErrorMsg = ex.Message;
                LogUtils.WriteLog("Step 2: Chọn Mục đích lưu trú gặp lỗi : " + ex.StackTrace, "err");
                ShowMessage(this, "Step 2: Chọn Mục đích lưu trú gặp lỗi : " + ex.Message);
                WriteLog("Step 2: Chọn Mục đích lưu trú gặp lỗi : " + ex.Message);
            }
        }

        // Step 3: Nhập mã captcha
        public void Step3_SendCaptcha()
        {
            try
            {
                _item.RequestTime = DateTime.Now;

                ShowMessage(this, string.Format("Step 3: Đang gửi mã captcha [{0}]...", _item.CaptchaText));
                WriteLog(string.Format("Step 3: Đang gửi mã captcha [{0}]...", _item.CaptchaText));
                _lastResponse = RequestAppointVisa(_lastResponse);

                DecaptchaFinish(this, !_lastResponse.Contains("cp1_lblCaptchaError"));

                if (_lastResponse.Contains("cp1_lblCaptchaError"))
                {
                    _item.Status = VarState.InvalidCaptcha;
                    _item.ErrorMsg = string.Format("Captcha: {0} ", _item.CaptchaText);
                    ShowMessage(this, "Step 3: Mã captcha không chính xác.");
                    WriteLog("Step 3: Mã captcha không chính xác. " + _item.CaptchaText);
                }
                else
                {
                    if (_lastResponse.Contains("cp1_lblNoDates"))
                    {
                        string rgInfo = @"<span[^>]*id=""cp1_lblNoDatesEmbassyInfo""[^>]*><br/>(?<value>(?<date>\d+/\d+/\d+)\s+(?<time>\d+:\d+:\d+)\s+(?<part>\w{2})[^<]*)</span>";
                        string noDate = Regex.Match(_lastResponse, "<span[^>]*id=\"cp1_lblNoDates\"[^>]*>(?<value>[^<]*)</span>").Groups["value"].ToString();
                        string noDatesEmbassy = Regex.Match(_lastResponse, rgInfo).Groups["value"].ToString();

                        string[] dateStr = Regex.Match(_lastResponse, rgInfo).Groups["date"].ToString().Split('/');
                        string[] timeStr = Regex.Match(_lastResponse, rgInfo).Groups["time"].ToString().Split(':');
                        string part = Regex.Match(_lastResponse, rgInfo).Groups["part"].ToString();

                        int[] date = new int[3];
                        for (int i = 0; i < dateStr.Length; i++)
                            date[i] = int.Parse(dateStr[i]);

                        int[] time = new int[3];
                        for (int i = 0; i < timeStr.Length; i++)
                            time[i] = int.Parse(timeStr[i]);

                        DateTime serverTime = new DateTime(date[2], date[0], date[1], time[0], time[1], time[2]);
                        if (part == "PM")
                            serverTime.AddHours(12);

                        _item.NotOpenTime = serverTime.ToString("dd/MM/yyyy hh:mm:ss ") + part;
                        _item.Status = VarState.NotOpen;
                        _item.ErrorMsg = noDatesEmbassy;
                        ShowMessage(this, string.Format("Step 3: Not Open: {0}", noDate));
                        WriteLog(string.Format("Step 3: Not Open: {0}", noDate));
                    }
                    else
                    {
                        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                        doc.LoadHtml(_lastResponse);

                        HtmlNodeCollection nodes = doc.DocumentNode.SelectNodes("//table[@id='cp1_rblDate']//input[@type='radio']");
                        if (nodes != null)
                        {
                            int count = 0;
                            _appointDate.Clear();
                            foreach (HtmlNode node in nodes)
                            {
                                if (count > 6) break;
                                KeyValueItem item = new KeyValueItem(node.Attributes["value"].Value, node.ParentNode.InnerText.Replace("&nbsp;", " "));
                                _appointDate.Add(item);
                            }
                            ShowMessage(this, "Step 3: Xác nhận mã captcha thành công.");
                            WriteLog("Step 3: Xác nhận mã captcha thành công.");

                            /*
                             * Update 10.07.2015 by Shanks Tùng
                             * Gộp 3 bước: step 4, 5, 6 làm một
                             * Step 4 => Step 7 thay vì Step 4 => Step 5
                             * Các postData được gộp lại với nhau
                             * Đã test: OK
                             */
                            Step4_AppointDate2();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _item.Status = VarState.Error;
                _item.ErrorMsg = ex.Message;
                LogUtils.WriteLog("Step 3: Gửi captcha lỗi : " + ex.StackTrace, "err");
                ShowMessage(this, "Step 3: Gửi captcha lỗi : " + ex.Message);
                WriteLog("Step 3: Gửi captcha lỗi : " + ex.Message);
            }
        }

        // Step 4: Chọn ngày giờ cuộc hẹn
        public void Step4_AppointDate()
        {
            try
            {
                KeyValueItem selectedItem = null;
                int randomIndex = _rand.Next(0, _appointDate.Count);
                if (_appointDate.Count > 0)
                {
                    selectedItem = _appointDate[randomIndex];
                }

                if (selectedItem == null)
                {
                    ShowMessage(this, "Step 4: Không tìm thấy cuộc hẹn nào");
                }

                ShowMessage(this, string.Format("Step 4: Đang chọn ngày hẹn: {0}...", selectedItem.Value));
                _lastResponse = RequestAppointDateTime(_lastResponse, selectedItem.Key);
                ShowMessage(this, "Step 4: Đã chọn xong ngày hẹn: " + selectedItem.Value);

                _item.OpenTime = string.Format("[{0}]{1}", randomIndex, selectedItem.Value);
                Step5_SendPersonalInfo();
            }
            catch (Exception ex)
            {
                _item.Status = VarState.Error;
                _item.ErrorMsg = ex.Message;
                LogUtils.WriteLog("Step 4: Chọn ngày hẹn gặp lỗi : " + ex.StackTrace, "err");
                ShowMessage(this, "Step 4: Chọn ngày hẹn gặp lỗi : " + ex.Message);
            }
        }

        // Step 5: Gửi thông tin cá nhân
        public void Step5_SendPersonalInfo()
        {
            try
            {
                ShowMessage(this, "Step 5: Đang gửi thông tin cá nhân...");
                _lastResponse = RequestPersonalInfo(_lastResponse);
                int retry = 0;
                while (_lastResponse.Contains("Personal information") && retry++ < 3)
                {
                    ShowMessage(this, string.Format("Step 4: Đang thử lại lần [{0}]...", retry));
                    _lastResponse = RequestPersonalInfo(_lastResponse);
                }

                if (_lastResponse.Contains("Personal information"))
                {
                    ShowMessage(this, "Step 5: Không gửi được Thông tin cá nhân");
                    return;
                }

                ShowMessage(this, "Step 5: Đã gửi thông tin cá nhân");

                Step6_SendPassportContact();
            }
            catch (Exception ex)
            {
                _item.Status = VarState.Error;
                _item.ErrorMsg = ex.Message;
                LogUtils.WriteLog("Step 5: Gửi thông tin cá nhân gặp lỗi : " + ex.StackTrace, "err");
                ShowMessage(this, "Step 5: Gửi thông tin cá nhân gặp lỗi : " + ex.Message);
            }
        }

        public string _resultHtml = string.Empty;
        // Step 6: Gửi thông tin liên lạc và hộ chiếu
        public void Step6_SendPassportContact()
        {
            try
            {
                ShowMessage(this, "Step 6: Đang gửi thông tin liên lạc và hộ chiếu...");
                _lastResponse = RequestPassportContact(_lastResponse);

                int retry = 0;
                while (_lastResponse.Contains("Passport and contact") && retry++ < 3)
                {
                    ShowMessage(this, string.Format("Passport and contact: đang thử lại lần [{0}]...", retry));
                    _lastResponse = RequestPassportContact(_lastResponse);
                }

                if (_lastResponse.Contains("Passport and contact"))
                {
                    ShowMessage(this, "Step 6: Không gửi được Thông tin liên lạc và hộ chiếu");
                    return;
                }
                ShowMessage(this, "Step 6: Đã gửi thông tin liên lạc và hộ chiếu");
                _resultHtml = _lastResponse;

                Step7_PreviewInfo();
            }
            catch (Exception ex)
            {
                _item.Status = VarState.Error;
                _item.ErrorMsg = ex.Message;
                LogUtils.WriteLog("Step 6: Gửi thông tin liên lạc và hộ chiếu gặp lỗi : " + ex.StackTrace, "err");
                ShowMessage(this, "Step 6: Gửi thông tin liên lạc và hộ chiếu gặp lỗi : " + ex.Message);
            }
        }

        // Step 7: Xác nhận thông tin đăng ký
        public void Step7_PreviewInfo()
        {
            try
            {
                //ShowMessage(this, "Step 7: Đang xác nhận thông tin đăng ký...");
                //_lastResponse = RequestSubmit(_lastResponse);

                int retry = 0;
                string res = _lastResponse;
                bool success = false;
                while (!success && retry++ < Setting.RetrySubmitCount)
                {
                    Thread.Sleep(Setting.DelaySubmit);
                    ShowMessage(this, string.Format("Step 7: Xác nhận thông tin đăng ký lần thứ [{0}]...", retry));
                    WriteLog(string.Format("Step 7: Xác nhận thông tin đăng ký lần thứ [{0}]...", retry));

                    Parallel.For(0, Setting.ForceSubmitCount, (i) =>
                    {
                        Thread.Sleep(Setting.ForceSubmitDelay * i);
                        bool suc = false;
                        ShowMessage(this, string.Format("Step 7: Force Submit [{0}]...", i));

                        WriteLog(string.Format("Step 7: Start Force Submit [{0}]...", i));

                        string resHtml = RequestSubmit(res, out suc);

                        WriteLog(string.Format("Step 7: Finish Force Submit [{0}]. Result: {1}", i, suc ? "OK" : "Failed"));

                        if (suc)
                        {
                            success = true;
                            _lastResponse = resHtml;
                        }
                    });
                }

                _item.ResponseTime = DateTime.Now;
                long duration = (long)(_item.ResponseTime - _item.RequestTime).TotalMilliseconds;

                if (success)
                {
                    _item.Status = VarState.Success;
                    _item.ErrorMsg = string.Empty;
                    ShowMessage(this, string.Format("Step 7: Submit OK [{0}]", duration));
                    WriteLog(string.Format("Step 7: Submit OK [{0}]", duration));
                }
                else
                {
                    string error = Regex.Match(_lastResponse, "<span id=\"cp1_txtProblem\">(?<value>[^<]*)</span>").Groups["value"].ToString().Trim();
                    if (string.IsNullOrEmpty(error))
                    {
                        _item.Status = VarState.NotSubmit;
                        _item.ErrorMsg = "Server return Empty";
                        ShowMessage(this, string.Format("Step 7: Submit Failed [{0}]! ErrorMsg: Not Submit !", duration));
                        WriteLog(string.Format("Step 7: Submit Failed [{0}]! ErrorMsg: Not Submit !", duration));
                    }
                    else
                    {
                        _item.Status = VarState.Failed;
                        _item.ErrorMsg = error;
                        ShowMessage(this, string.Format("Step 7: Submit Failed [{0}ms] ! ErrorMsg: {1}", duration, error));
                        WriteLog(string.Format("Step 7: Submit Failed [{0}ms] ! ErrorMsg: {1}", duration, error));
                    }
                }
            }
            catch (Exception ex)
            {
                _item.Status = VarState.Error;
                _item.ErrorMsg = ex.Message;
                LogUtils.WriteLog("Step 7: Xác nhận thông tin đăng ký gặp lỗi : " + ex.StackTrace, "err");
                ShowMessage(this, "Step 7: Xác nhận thông tin đăng ký gặp lỗi : " + ex.Message);
                WriteLog("Step 7: Xác nhận thông tin đăng ký gặp lỗi : " + ex.Message);
            }
        }

        public void PreviousStep2()
        {
            try
            {
                _item.Status = VarState.Running;
                _lastRun = DateTime.Now;
                ShowMessage(this, "Đang quay lại bước nhập captcha...");
                WriteLog("Quay lại bước nhập captcha...");
                _lastResponse = RequestPreviousStep2(_lastResponse);

                if (http.LastUrl.Contains("disclaimer"))
                    Step1_SelectEmbassy();
                else
                {
                    RefreshCaptcha(_lastResponse);
                    //ShowMessage(this, "Đã quay lại bước nhập captcha !");
                }
            }
            catch (Exception ex)
            {
                _item.Status = VarState.Error;
                _item.ErrorMsg = ex.Message;
                LogUtils.WriteLog("Quay lại bước nhập captcha gặp lỗi : " + ex.StackTrace, "err");
                ShowMessage(this, "Quay lại bước nhập captcha gặp lỗi : " + ex.Message);
                WriteLog("Quay lại bước nhập captcha gặp lỗi : " + ex.Message);
            }
        }

        #region fields for request
        private HttpHelper http = new HttpHelper();

        private string _lastCaptchaUrl = string.Empty;

        private string _lastResponse = string.Empty;
        public string LastResponse
        {
            get { return _lastResponse; }
            set { _lastResponse = value; }
        }

        private string _rgEventState = "<input type=\"hidden\" name=\"__EVENTTARGET\" id=\"__EVENTTARGET\" value=\"(?<value>.*?)\" />";
        private string _rgViewState = "<input type=\"hidden\" name=\"__VIEWSTATE\" id=\"__VIEWSTATE\" value=\"(?<value>.*?)\" />";
        private string _rgEventValidation = "<input type=\"hidden\" name=\"__EVENTVALIDATION\" id=\"__EVENTVALIDATION\" value=\"(?<value>.*?)\" />";
        private string _rgUlList = "<ul class=\"rcbList\"[^>]*>[^$]*</ul>";

        private List<KeyValueItem> _listCitizenship = new List<KeyValueItem>();
        private List<KeyValueItem> _listResidence = new List<KeyValueItem>();
        private List<KeyValueItem> _listEmbassy = new List<KeyValueItem>();
        private List<KeyValueItem> _listPurposeOfStay = new List<KeyValueItem>();
        #endregion

        #region Request
        private string RequestCitizenship(string content)
        {
            int selectedIndex = _listPurposeOfStay.IndexOf(_item.PurposeOfStay);

            string rsm1 = GetRsm1(content);
            string viewState = GetViewState(content);
            string eventValidation = GetEventValidation(content);
            string eventArgument = EncodeDataString(string.Format("{{\"Command\":\"Select\",\"Index\":{0}}}", selectedIndex));
            string clientState = EncodeDataString(string.Format("{{\"logEntries\":[],\"value\":\"{0}\",\"text\":\"{1}\",\"enabled\":true,\"checkedIndices\":[],\"checkedItemsTextOverflows\":false}}", _item.Citizenship.Key, _item.Citizenship.Value));

            string postData = string.Format("rsm1_TSM={0}&__EVENTTARGET=ctl00%24cp1%24ddCitizenship&__EVENTARGUMENT={3}&__VIEWSTATE={1}&__VIEWSTATEENCRYPTED=&__EVENTVALIDATION={2}"
                                        + "&ctl00%24ddLocale=English+%28United+Kingdom%29"
                                        + "&ctl00_ddLocale_ClientState="
                                        + "&ctl00%24cp1%24ddCitizenship={4}"
                                        + "&ctl00_cp1_ddCitizenship_ClientState={5}"
                                        + "&ctl00%24cp1%24ddCountryOfResidence=Select+your+country+of+residence"
                                        + "&ctl00_cp1_ddCountryOfResidence_ClientState="
                                        + "&ctl00%24cp1%24ddEmbassy=Please+select+the+Czech+Embassy%2FConsulate+you+wish+to+visit."
                                        + "&ctl00_cp1_ddEmbassy_ClientState="
                                        + "&ctl00%24cp1%24ddVisaType=Choose+your+purpose+of+stay"
                                        + "&ctl00_cp1_ddVisaType_ClientState="
                                        + "&ctl00_cp1_btnNext_ClientState="
                                        , rsm1
                                        , viewState
                                        , eventValidation
                                        , eventArgument
                                        , EncodeDataString(_item.Citizenship.Value)
                                        , clientState);

            return http.FetchHttpPost("https://visapoint.eu/form", http.LastUrl, postData);
        }

        private string RequestPurposeOfStay(string content)
        {
            KeyValueItem itemEmbassy = _item.Embassy;

            int selectedIndex = _listPurposeOfStay.IndexOf(_item.PurposeOfStay);

            string rsm1 = GetRsm1(content);
            string viewState = GetViewState(content);
            string eventValidation = GetEventValidation(content);
            string eventArgument = EncodeDataString(string.Format("{{\"Command\":\"Select\",\"Index\":{0}}}", selectedIndex));
            string clientState = EncodeDataString(string.Format("{{\"logEntries\":[],\"value\":\"{0}\",\"text\":\"{1}\",\"enabled\":true,\"checkedIndices\":[],\"checkedItemsTextOverflows\":false}}", _item.PurposeOfStay.Key, _item.PurposeOfStay.Value));

            string postData = string.Format("rsm1_TSM={0}"
                                        + "&__EVENTTARGET=ctl00%24cp1%24ddVisaType"
                                        + "&__EVENTARGUMENT={3}"
                                        + "&__VIEWSTATE={1}"
                                        + "&__VIEWSTATEENCRYPTED="
                                        + "&__EVENTVALIDATION={2}"
                                        + "&ctl00%24ddLocale=English+%28United+Kingdom%29"
                                        + "&ctl00_ddLocale_ClientState="
                                        + "&ctl00%24cp1%24ddCitizenship={4}"
                                        + "&ctl00_cp1_ddCitizenship_ClientState="
                                        + "&ctl00%24cp1%24ddCountryOfResidence={5}"
                                        + "&ctl00_cp1_ddCountryOfResidence_ClientState="
                                        + "&ctl00%24cp1%24ddEmbassy={6}"
                                        + "&ctl00_cp1_ddEmbassy_ClientState="
                                        + "&ctl00%24cp1%24ddVisaType={7}"
                                        + "&ctl00_cp1_ddVisaType_ClientState={8}"
                                        + "&ctl00_cp1_btnNext_ClientState="
                                        , rsm1
                                        , viewState
                                        , eventValidation
                                        , eventArgument
                                        , EncodeDataString(_item.Citizenship.Value)
                                        , EncodeDataString(_item.Citizenship.Value)
                                        , EncodeDataString(itemEmbassy.Value)
                                        , EncodeDataString(_item.PurposeOfStay.Value)
                                        , clientState);

            return http.FetchHttpPost("https://visapoint.eu/form", http.LastUrl, postData);
        }

        private void RefreshCaptcha(string content = "")
        {
            Bitmap bmp = null;
            if (!string.IsNullOrEmpty(content))
            {
                string rgCaptcha = "<img[^>]*id=\"c_pages_form_cp1_captcha1_CaptchaImage\" [^>]*src=\"(?<value>[^\"]+)\"[^>]*>";
                Match m = Regex.Match(content, rgCaptcha);
                if (m.Success)
                {
                    string imgUrl = "https://visapoint.eu" + m.Groups["value"].ToString();
                    bmp = http.FetchHttpImage(imgUrl, http.LastUrl);
                    _lastCaptchaUrl = imgUrl;
                    Decaptcha(bmp);
                }
                else
                {
                    bmp = new Bitmap(250, 50);
                }
            }
            else
            {
                if (Regex.IsMatch(_lastCaptchaUrl, "d=\\d+"))
                    _lastCaptchaUrl = Regex.Replace(_lastCaptchaUrl, "d=\\d+", "d=" + DateTime.Now.Ticks);
                else
                    _lastCaptchaUrl += "&d=" + DateTime.Now.Ticks;
                bmp = http.FetchHttpImage(_lastCaptchaUrl, "https://visapoint.eu/form");
                Decaptcha(bmp);
            }
            _item.CaptchaImg = bmp;
        }

        private string RequestAppointVisa(string content)
        {
            KeyValueItem itemEmbassy = _item.Embassy;

            int selectedIndex = _listPurposeOfStay.IndexOf(_item.PurposeOfStay);

            string rsm1 = GetRsm1(content);
            string viewState = GetViewState(content);
            string eventValidation = GetEventValidation(content);
            string clientState = EncodeDataString(string.Format("{{\"enabled\":true,\"emptyMessage\":\"\",\"validationText\":\"{0}\",\"valueAsString\":\"{0}\",\"valueWithPromptAndLiterals\":\"{0}\",\"lastSetTextBoxValue\":\"{0}\"}}", _item.CaptchaText));

            string t = Regex.Match(_lastCaptchaUrl, "t=(?<value>[^&]{10,})").Groups["value"].ToString();
            string wrapper = Regex.Match(content, "ctl00_cp1_(?<value>[^_]+)_wrapper").Groups["value"].ToString();

            string postData = string.Format("rsm1_TSM={0}"
                                        + "&__EVENTTARGET=ctl00%24cp1%24btnNext"
                                        + "&__EVENTARGUMENT={3}"
                                        + "&__VIEWSTATE={1}"
                                        + "&__VIEWSTATEENCRYPTED="
                                        + "&__EVENTVALIDATION={2}"
                                        + "&ctl00%24ddLocale=English+%28United+Kingdom%29"
                                        + "&ctl00_ddLocale_ClientState="
                                        + "&ctl00%24cp1%24ddCitizenship={4}"
                                        + "&ctl00_cp1_ddCitizenship_ClientState="
                                        + "&ctl00%24cp1%24ddCountryOfResidence={5}"
                                        + "&ctl00_cp1_ddCountryOfResidence_ClientState="
                                        + "&ctl00%24cp1%24ddEmbassy={6}"
                                        + "&ctl00_cp1_ddEmbassy_ClientState="
                                        + "&ctl00%24cp1%24ddVisaType={7}"
                                        + "&ctl00_cp1_ddVisaType_ClientState="
                                        + "&ctl00%24cp1%24{11}={8}"
                                        + "&ctl00_cp1_{11}_ClientState={9}"
                                        + "&LBD_VCID_c_pages_form_cp1_captcha1={10}"
                                        + "&ctl00_cp1_btnNext_ClientState="
                                        , rsm1
                                        , viewState
                                        , eventValidation
                                        , string.Empty
                                        , EncodeDataString(_item.Citizenship.Value)
                                        , EncodeDataString(_item.Citizenship.Value)
                                        , EncodeDataString(itemEmbassy.Value)
                                        , EncodeDataString(_item.PurposeOfStay.Value)
                                        , EncodeDataString(_item.CaptchaText)
                                        , clientState
                                        , t
                                        , wrapper);

            return http.FetchHttpPost("https://visapoint.eu/form", http.LastUrl, postData);

        }

        private string RequestAppointDateTime(string content, string value)
        {
            string rsm1 = GetRsm1(content);
            string viewState = GetViewState(content);
            string eventValidation = GetEventValidation(content);

            string postData = string.Format("rsm1_TSM={0}&__EVENTTARGET=ctl00%24cp1%24btnNext&__EVENTARGUMENT=&__VIEWSTATE={1}&__VIEWSTATEENCRYPTED=&__EVENTVALIDATION={2}&ctl00%24ddLocale=English+%28United+Kingdom%29&ctl00_ddLocale_ClientState=&ctl00%24cp1%24rblDate={3}&ctl00_cp1_btnPrev_ClientState=&ctl00_cp1_btnNext_ClientState=", rsm1, viewState, eventValidation, value);

            return http.FetchHttpPost("https://visapoint.eu/form", http.LastUrl, postData);
        }

        private string RequestPersonalInfo(string content)
        {
            string rsm1 = GetRsm1(content);
            string viewState = GetViewState(content);
            string eventValidation = GetEventValidation(content);

            string firstName = _item.Name;
            string firstName_ClientState = EncodeDataString(string.Format("{{\"enabled\":true,\"emptyMessage\":\"Fill in your first name(s)\",\"validationText\":\"{0}\",\"valueAsString\":\"{0}\",\"lastSetTextBoxValue\":\"{0}\"}}", firstName));

            string familyName = _item.FamilyName;
            string familyName_ClientState = EncodeDataString(string.Format("{{\"enabled\":true,\"emptyMessage\":\"Fill in your Family Name(s)\",\"validationText\":\"{0}\",\"valueAsString\":\"{0}\",\"lastSetTextBoxValue\":\"{0}\"}}", familyName));

            DateTime birthDate = DateTime.ParseExact(_item.DateOfBirth, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            string birthDate_input_ClientState = EncodeDataString(string.Format("{{\"enabled\":true,\"emptyMessage\":\"\",\"validationText\":\"{0}-00-00-00\",\"valueAsString\":\"{0}-00-00-00\",\"minDateStr\":\"{0}-00-00-00\",\"maxDateStr\":\"{2}-00-00-00\",\"lastSetTextBoxValue\":\"{1}\"}}", birthDate.ToString("yyyy-MM-dd"), birthDate.ToString("dd/MM/yyyy"), DateTime.Now.ToString("yyyy-MM-dd")));

            string birthDate_ClientState = EncodeDataString(string.Format("{{\"minDateStr\":\"{0}-00-00-00\",\"maxDateStr\":\"{1}-00-00-00\"}}", birthDate.ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd")));

            string calender_sd = EncodeDataString("[]");
            string calender_ad = EncodeDataString(string.Format("[[{0},{1}], [{2}], [{2}]]", birthDate.Year, DateTime.Now.ToString("M,d"), DateTime.Now.ToString("yyyy,M,d")));

            string countryCode = _item.CountryOfBirth.Key;
            string birthCountry = _item.CountryOfBirth.Value;

            string birthCountry_ClientState = EncodeDataString(string.Format("{{\"logEntries\":[],\"value\":\"{0}\",\"text\":\"{1}\",\"enabled\":true,\"checkedIndices\":[],\"checkedItemsTextOverflows\":false}}", countryCode, birthCountry));

            string sex = _item.Gender;
            string sex_ClientState = EncodeDataString(string.Format("{{\"logEntries\":[],\"value\":\"F\",\"text\":\"{0}\",\"enabled\":true,\"checkedIndices\":[],\"checkedItemsTextOverflows\":false}}", sex));

            string postData = string.Format("rsm1_TSM={0}"
                                        + "&__LASTFOCUS="
                                        + "&__EVENTTARGET=ctl00%24cp1%24btnNext"
                                        + "&__EVENTARGUMENT="
                                        + "&__VIEWSTATE={1}"
                                        + "&__VIEWSTATEENCRYPTED="
                                        + "&__EVENTVALIDATION={2}"
                                        + "&ctl00%24ddLocale=English+%28United+Kingdom%29"
                                        + "&ctl00_ddLocale_ClientState="
                                        + "&ctl00%24cp1%24txtFirstName={3}"
                                        + "&ctl00_cp1_txtFirstName_ClientState={4}"
                                        + "&ctl00%24cp1%24txtFamilyName={5}"
                                        + "&ctl00_cp1_txtFamilyName_ClientState={6}"
                                        + "&ctl00%24cp1%24txtBirthDate={7}"
                                        + "&ctl00%24cp1%24txtBirthDate%24dateInput={8}"
                                        + "&ctl00_cp1_txtBirthDate_dateInput_ClientState={9}"
                                        + "&ctl00_cp1_txtBirthDate_calendar_SD={10}"
                                        + "&ctl00_cp1_txtBirthDate_calendar_AD={11}"
                                        + "&ctl00_cp1_txtBirthDate_ClientState={12}"
                                        + "&ctl00%24cp1%24ddBirthCountry={13}"
                                        + "&ctl00_cp1_ddBirthCountry_ClientState={14}"
                                        + "&ctl00%24cp1%24ddSex={15}"
                                        + "&ctl00_cp1_ddSex_ClientState={16}"
                                        + "&ctl00_cp1_btnPrev_ClientState="
                                        + "&ctl00_cp1_btnNext_ClientState="
                                        , rsm1, viewState, eventValidation
                                        , EncodeDataString(firstName), firstName_ClientState
                                        , EncodeDataString(familyName), familyName_ClientState
                                        , EncodeDataString(birthDate.ToString("yyyy-MM-dd"))
                                        , EncodeDataString(birthDate.ToString("dd/MM/yyyy"))
                                        , birthDate_input_ClientState
                                        , calender_sd, calender_ad
                                        , birthDate_ClientState
                                        , birthCountry
                                        , birthCountry_ClientState
                                        , EncodeDataString(sex), sex_ClientState);

            return http.FetchHttpPost("https://visapoint.eu/form", http.LastUrl, postData);
        }

        private string RequestPassportContact(string content)
        {
            string rsm1 = GetRsm1(content);
            string viewState = GetViewState(content);
            string eventValidation = GetEventValidation(content);

            string passport = _item.Passport;
            string passport_ClientState = EncodeDataString(string.Format("{{\"enabled\":true,\"emptyMessage\":\"Fill in your passport number\",\"validationText\":\"{0}\",\"valueAsString\":\"{0}\",\"lastSetTextBoxValue\":\"{0}\"}}", passport));

            string email = _item.Email;
            string email_ClientState = EncodeDataString(string.Format("{{\"enabled\":true,\"emptyMessage\":\"Fill in valid email\",\"validationText\":\"{0}\",\"valueAsString\":\"{0}\",\"lastSetTextBoxValue\":\"{0}\"}}", email));

            string phone = _item.Phone;
            string phone_ClientState = EncodeDataString(string.Format("{{\"enabled\":true,\"emptyMessage\":\"Fill in your phone number\",\"validationText\":\"{0}\",\"valueAsString\":\"{0}\",\"valueWithPromptAndLiterals\":\"{0}\",\"lastSetTextBoxValue\":\"{0}\"}}", phone));

            string postData = string.Format("rsm1_TSM={0}"
                                        + "&__LASTFOCUS="
                                        + "&__EVENTTARGET=ctl00%24cp1%24btnNext"
                                        + "&__EVENTARGUMENT="
                                        + "&__VIEWSTATE={1}"
                                        + "&__VIEWSTATEENCRYPTED="
                                        + "&__EVENTVALIDATION={2}"
                                        + "&ctl00%24ddLocale=English+%28United+Kingdom%29"
                                        + "&ctl00_ddLocale_ClientState="
                                        + "&ctl00%24cp1%24txtPassportNumber={3}"
                                        + "&ctl00_cp1_txtPassportNumber_ClientState={4}"
                                        + "&ctl00%24cp1%24txtEmail={5}"
                                        + "&ctl00_cp1_txtEmail_ClientState={6}"
                                        + "&ctl00%24cp1%24txtPhone={7}"
                                        + "&ctl00_cp1_txtPhone_ClientState={8}"
                                        + "&ctl00_cp1_btnPrev_ClientState="
                                        + "&ctl00_cp1_btnNext_ClientState="
                                        , rsm1, viewState, eventValidation
                                        , EncodeDataString(passport), passport_ClientState
                                        , EncodeDataString(email), email_ClientState
                                        , EncodeDataString(phone), phone_ClientState);

            return http.FetchHttpPost("https://visapoint.eu/form", http.LastUrl, postData);
        }

        private string RequestSubmit(string content, out bool success)
        {
            string rsm1 = GetRsm1(content);
            string viewState = GetViewState(content);
            string eventValidation = GetEventValidation(content);

            string postData = string.Format("rsm1_TSM={0}"
                                        + "&__EVENTTARGET=ctl00%24cp1%24btnSend"
                                        + "&__EVENTARGUMENT="
                                        + "&__VIEWSTATE={1}"
                                        + "&__VIEWSTATEENCRYPTED="
                                        + "&__EVENTVALIDATION={2}"
                                        + "&ctl00%24ddLocale=English+%28United+Kingdom%29"
                                        + "&ctl00_ddLocale_ClientState="
                                        + "&ctl00_cp1_btnPrev_ClientState="
                                        + "&ctl00_cp1_btnSend_ClientState="
                                        , rsm1
                                        , viewState
                                        , eventValidation);
            return http.FetchHttpPostSubmit("https://visapoint.eu/form", http.LastUrl, postData, out success);
        }

        private string RequestPreviousStep2(string content)
        {
            string rsm1 = GetRsm1(content);
            string viewState = GetViewState(content);
            string eventValidation = GetEventValidation(content);

            string t = Regex.Match(_lastCaptchaUrl, "t=(?<value>[^&]{10,})").Groups["value"].ToString();

            string postData = string.Format("rsm1_TSM={0}"
                                            + "&__EVENTTARGET=ctl00%24cp1%24btnPrev"
                                            + "&__EVENTARGUMENT="
                                            + "&__VIEWSTATE={1}"
                                            + "&__VIEWSTATEENCRYPTED="
                                            + "&__EVENTVALIDATION={2}"
                                            + "&ctl00%24ddLocale=English+%28United+Kingdom%29"
                                            + "&ctl00_ddLocale_ClientState="
                                            + "&ctl00%24cp1%24rblDate={3}"
                                            + "&ctl00_cp1_btnPrev_ClientState="
                                            + "&ctl00_cp1_btnNext_ClientState="
                                            , rsm1, viewState, eventValidation, t);

            return http.FetchHttpPost("https://visapoint.eu/form", http.LastUrl, postData);
        }
        #endregion

        #endregion

        #region Utils

        private void Decaptcha(Bitmap bmp)
        {
            ShowMessage(this, string.Format("Decaptcha..."));
            WriteLog(string.Format("Decaptcha..."));
            string captchaText = string.Empty;

            int providerIndex = _captchaProviderIndex;
            if (Setting.CaptchaManual == 1 && _countManual < Setting.MaxCaptchaManual)
            {
                providerIndex = -1;
            }

            switch (providerIndex)
            {
                case -1:
                    _startManualCaptcha(this, bmp);
                    captchaText = CaptchaManual;
                    break;
                case 0:
                    captchaText = DeathByCaptchaDecaptcher.Decaptcha(bmp,
                                                                    Setting.UsernameDbc,
                                                                    Setting.PasswordDbc,
                                                                    Setting.DecaptchaTimeout);
                    break;
                case 1:
                    captchaText = DeCaptcherDecaptcher.Decaptcha(bmp,
                                                                    Setting.UsernameDeCaptcher,
                                                                    Setting.PasswordDeCaptcher,
                                                                    Setting.DecaptchaTimeout);
                    break;
                case 2:
                    captchaText = ImageTyperzDecaptcher.Decaptcha(bmp,
                                                                Setting.UsernameImagetyperz,
                                                                Setting.PasswordImagetyperz,
                                                                Setting.DecaptchaTimeout);
                    break;
                case 3:
                    captchaText = ShanibpoDecaptcher.Decaptcha(bmp,
                                                                Setting.UsernameShanibpo,
                                                                Setting.PasswordShanibpo,
                                                                Setting.DecaptchaTimeout);
                    break;
            }

            _item.CaptchaText = RefineCaptchaString(captchaText);
            ShowMessage(this, string.Format("Decaptcha thành công: " + _item.CaptchaText));
            WriteLog(string.Format("Decaptcha thành công: " + _item.CaptchaText));
        }

        public static byte[] ImageToByte(Image img)
        {
            ImageConverter converter = new ImageConverter();
            return (byte[])converter.ConvertTo(img, typeof(byte[]));
        }

        private string RefineCaptchaString(string captcha)
        {
            string text = captcha.Trim();
            string s = string.Empty;
            for (int i = 1; i < text.Length; i++)
                if (text[i] != ' ' && text[i - 1] != ' ')
                    s += text[i - 1].ToString() + ' ';
                else
                    s += text[i - 1];
            if (text.Length > 0)
                s += text[text.Length - 1];
            s = s.ToUpper();
            return s;
        }

        #region Country KeyValue
        public static List<string> ListCountryValue = new List<string> { "Albania (Shqipëria)", "Azerbaijan (Azərbaycanca)", "Belarus (Беларусь)", "Egypt (مصر)", "Georgia (საქართველო)", "Indonesia (Indonesia)", "Jordan (الأردن)", "Kazakhstan (Қазақстан)", "Libya (ليبيا)", "Moldova (Moldova)", "Mongolia (Монгол&nbsp;улс)", "Nigeria (Nigeria)", "People's Republic of China (中华人民共和国)", "Republic of the Philippines (Philippines)", "Russia (Россия)", "Serbia (Srbija)", "Thailand (ไทย)", "Turkey (Türkiye)", "U.A.E. (الإمارات العربية المتحدة)", "Ukraine (Україна)", "Uzbekistan (U'zbekiston Respublikasi)", "Vietnam (Việt Nam)", "Afghanistan (Afghanistan)", "Algeria (الجزائر)", "American Samoa (American Samoa)", "Andorra (Andorra)", "Angola (Angola)", "Anguilla (Anguilla)", "Antarctica (Antarctica)", "Antigua (Antigua)", "Argentina (Argentina)", "Armenia (Հայաստան)", "Aruba (Aruba)", "Australia (Australia)", "Austria (Österreich)", "Bahamas (Bahamas)", "Bahrain (البحرين)", "Bangladesh (Bangladesh)", "Barbados (Barbados)", "Belgium (Belgique)", "Belize (Belize)", "Benin (Benin)", "Bermuda (Bermuda)", "Bhutan (Bhutan)", "Bolivia (Bolivia)", "Bosnia and Herzegovina (Босна и Херцеговина)", "Botswana (Botswana)", "Bouvet Island (Bouvet Island)", "Brazil (Brasil)", "British Indian Ocean Territory (British Indian Ocean Territory)", "British Virgin Islands (British Virgin Islands)", "Brunei Darussalam (Brunei Darussalam)", "Bulgaria (България)", "Burkina Faso (UVolta) (Burkina Faso (UVolta))", "Burma (Burma)", "Burundi (Burundi)", "Cambodia (Cambodia)", "Cameroon (Cameroon)", "Canada (Canada)", "Cape Verde (Cape Verde)", "Caribbean (Caribbean)", "Cayman Islands (Cayman Islands)", "Central African Republic (Central African Republic)", "Cocos (Keeling) Islands (Cocos (Keeling) Islands)", "Colombia (Colombia)", "Comoro Islands (Comoro Islands)", "Congo (Congo)", "Cook Islands (Cook Islands)", "Costa Rica (Costa Rica)", "Croatia (Hrvatska)", "Cuba (Cuba)", "Cyprus (Cyprus)", "Czech Republic (Česká&nbsp;republika)", "Dem Rep Congo (Dem Rep Congo)", "Denmark (Danmark)", "Djibouti (Djibouti)", "Dominica (Dominica)", "Dominican Republic (República Dominicana)", "Ecuador (Ecuador)", "El Salvador (El Salvador)", "Equat Guinea (Equat Guinea )", "Eritrea (Eritrea)", "Estonia (Eesti)", "Ethiopia (Ethiopia)", "Falkland Islands (Islas Malvinas) (Falkland Islands (Islas Malvinas))", "Faroe Islands (Føroyar)", "Fiji (Fiji)", "Finland (Suomi)", "France (France)", "French Guiana (French Guiana)", "French Polynesia (French Polynesia)", "French Southern and Antarctic Lands (French Southern and Antarctic Lands)", "Gabon (Gabon)", "Germany (Deutschland)", "Ghana (Ghana)", "Gibraltar (Gibraltar)", "Greece (Ελλάδα)", "Greenland (Greenland)", "Grenada (Grenada)", "Guadeloupe (Guadeloupe)", "Guam (Guam)", "Guatemala (Guatemala)", "Guinea (Guinea)", "Guinea-Bissau (Guinea-Bissau)", "Guyana (Guyana)", "Haiti (Haiti)", "Heard Islands and McDonald Islands (Heard Islands and McDonald Islands)", "Honduras (Honduras)", "Hong Kong S.A.R. (香港特别行政區)", "Hungary (Magyarország)", "Chad (Chad)", "Chile (Chile)", "Christmas Island (Christmas Island)", "Iceland (Ísland)", "India (भारत)", "Iran (ايران)", "Iraq (العراق)", "Ireland (Eire)", "Israel (ישראל)", "Italy (Italia)", "Ivory Coast (Ivory Coast)", "Jamaica (Jamaica)", "Japan (日本)", "Kenya (Kenya)", "Kiribati (Kiribati)", "Korea (대한민국)", "Kosovo (Kosovo)", "Kuwait (الكويت)", "Kyrgyzstan (Кыргызстан)", "Laos (Laos)", "Latvia (Latvija)", "Lebanon (لبنان)", "Lesotho (Lesotho)", "Liberia (Liberia)", "Liechtenstein (Liechtenstein)", "Lithuania (Lietuva)", "Luxembourg (Luxemburg)", "Macao S.A.R. (澳門特别行政區)", "Macau (Macau)", "Macedonia (FYROM) (Македонија)", "Madagascar (Malagasy Republic) (Madagascar (Malagasy Republic))", "Malawi (Malawi)", "Malaysia (Malaysia)", "Maldives (ދިވެހި ރާއްޖެ)", "Mali (Mali)", "Malta (Malta)", "Marshall Islands (Marshall Islands)", "Martinique (Martinique)", "Mauritania (Mauritania)", "Mauritius (Mauritius)", "Mayotte (Mayotte)", "Mexico (México)", "Micronesia (Micronesia)", "Montenegro (Montenegro)", "Montserrat (Montserrat)", "Morocco (المملكة المغربية)", "Mozambique (Mozambique)", "Namibia (Namibia)", "Nauru (Nauru)", "Nepal (Nepal)", "Netherlands (Nederland)", "New Caledonia (New Caledonia)", "New Zealand (New Zealand)", "Nicaragua (Nicaragua)", "Niger (Niger)", "Niue (Niue)", "Norfolk Island (Norfolk Island)", "Northern Mariana Islands (Northern Mariana Islands)", "Norway (Norge)", "Oman (عمان)", "Pakistan (پاکستان)", "Palau (Palau)", "Palestinia (Palestinia)", "Panama (Panamá)", "Papua New Guinea (Papua New Guinea)", "Paraguay (Paraguay)", "Peru (Perú)", "Pitcairn Islands (Pitcairn Islands)", "Poland (Polska)", "Portugal (Portugal)", "Principality of Monaco (Principauté de Monaco)", "Puerto Rico (Puerto Rico)", "Qatar (قطر)", "Reunion (Reunion)", "Romania (România)", "Rwanda (Rwanda)", "San Marino (San Marino)", "Sao Tome and Principe (Sao Tome and Principe)", "Saudi Arabia (المملكة العربية السعودية)", "Senegal (Senegal)", "Seychelles (Seychelles)", "Sierra Leone (Sierra Leone)", "Singapore (新加坡)", "Slovakia (Slovenská republika)", "Slovenia (Slovenija)", "Solomon Islands (Solomon Islands)", "Somalia (Somalia)", "South Africa (Suid Afrika)", "Spain (Espanya)", "Sri Lanka (Ceylon) (Sri Lanka (Ceylon))", "St. Helena (St. Helena)", "St. Kitts and Nevis (St. Kitts and Nevis)", "St. Lucia (St. Lucia)", "St. Pierre and Miquelon (St. Pierre and Miquelon)", "St. Vincent and The Grenadines (St. Vincent and The Grenadines)", "Sudan (Sudan)", "Suriname (Suriname)", "Svalbard and Jan Mayen (Svalbard and Jan Mayen)", "Swaziland (Swaziland)", "Sweden (Sverige)", "Switzerland (Schweiz)", "Syria (سوريا)", "Taiwan (台灣)", "Tajikistan (Tajikistan)", "Tanzania (Tanzania)", "The Gambia (The Gambia)", "Tibet (Tibet)", "Timor Leste (Timor Leste )", "Togo (Togo)", "Tokelau (Tokelau)", "Tonga (Tonga)", "Trinidad and Tobago (Trinidad y Tobago)", "Tunisia (تونس)", "Turkmenistan (Turkmenistan)", "Turks and Caicos Islands (Turks and Caicos Islands)", "Tuvalu (Tuvalu)", "Uganda (Uganda)", "United Kingdom (United Kingdom)", "United States (United States)", "United States Minor (United States Minor)", "Uruguay (Uruguay)", "Vanuatu (Vanuatu)", "Vatican City (Vatican City)", "Venezuela (Republica Bolivariana de Venezuela)", "Virgin Islands (Virgin Islands)", "Wallis and Futuna (Wallis and Futuna)", "Western Sahara (Western Sahara)", "Western Samoa (Western Samoa)", "Yemen (اليمن)", "Zaire (Zaire)", "Zambia (Zambia)", "Zimbabwe (Zimbabwe)" };

        public static List<int> ListCountryKey = new List<int> { 26, 41, 33, 64, 44, 31, 92, 48, 69, 186, 52, 196, 56, 99, 23, 62, 28, 29, 101, 32, 51, 39, 115, 73, 116, 117, 118, 119, 120, 121, 94, 40, 245, 67, 66, 123, 103, 124, 125, 60, 90, 126, 127, 128, 106, 111, 129, 130, 21, 248, 239, 63, 2, 131, 190, 132, 133, 134, 68, 135, 88, 136, 137, 140, 89, 141, 142, 144, 77, 24, 146, 147, 5, 143, 6, 148, 149, 83, 97, 107, 150, 151, 35, 152, 153, 45, 154, 10, 11, 155, 156, 157, 158, 7, 160, 161, 8, 162, 163, 164, 165, 72, 166, 167, 168, 169, 170, 108, 65, 13, 138, 100, 139, 14, 46, 38, 55, 79, 12, 15, 145, 85, 16, 50, 172, 17, 247, 98, 49, 173, 36, 95, 174, 175, 75, 37, 71, 74, 176, 42, 177, 178, 47, 54, 179, 112, 180, 181, 182, 183, 184, 59, 185, 187, 188, 78, 189, 191, 192, 193, 18, 194, 76, 109, 195, 197, 198, 199, 19, 84, 30, 200, 201, 80, 202, 104, 91, 203, 20, 61, 81, 110, 105, 204, 22, 205, 212, 213, 1, 214, 215, 216, 70, 25, 34, 217, 218, 43, 3, 219, 206, 207, 208, 209, 210, 220, 221, 222, 223, 27, 57, 53, 4, 224, 225, 159, 226, 227, 228, 229, 230, 93, 82, 231, 232, 233, 234, 58, 9, 246, 102, 238, 171, 86, 240, 241, 242, 211, 87, 243, 244, 96 };

        // Bản cũ
        //public static List<string> ListPurposeOfStayValue = new List<string> { "Long-term visa for Business", "Long-term visa for Other Purposes", "Application for Permanent Residence Permit", "Application for a long term residence for the purpose of employment \"Blue Card\"", "Long-term visa for purposes- health/culture/sport", "Long-term residence permit - family reunification", "Long-term residence permit for study", "Long-term residence permit - scientific research", "Long-term visa - family reunification", "Long-term visa for study", "Long-term visa – scientific research", "Transfer of a valid visa to a new travel document", "Long-term residence permit - Employment card" };

        //public static List<int> ListPurposeOfStayKey = new List<int> { 3, 6, 7, 10, 13, 14, 15, 16, 2, 4, 12, 9, 19 };

#if TESTMODE
        public static List<string> ListPurposeOfStayValue = new List<string> { 
                                                                "Application for a long term residence for the purpose of employment \"Blue Card\"",
                                                                "Application for Permanent Residence Permit",
                                                                "Long-term residence permit",
                                                                "Long-term residence permit - Employment card",
                                                                "Long-term residence permit - family reunification",
                                                                "Long-term residence permit - scientific research ",
                                                                "Long-term residence permit for study",
                                                                "Long-term visa",
                                                                "Short-term visa - Schengen" };

        public static List<int> ListPurposeOfStayKey = new List<int> { 10, 7, 20, 19, 14, 16, 15, 18, 11 };
#else
        // Bảm mới
        public static List<string> ListPurposeOfStayValue = new List<string> { "Application for Permanent Residence Permit", "Long-term residence permit", "Long-term visa", "Long-term visa for Business" };

        public static List<int> ListPurposeOfStayKey = new List<int> { 7, 20, 18, 3 };
#endif
        #endregion

        #region Parse Citizenship, Residence, Embassy
        private void ParseCitizenship(string content)
        {
            string rgCitizenship = @"__doPostBack\(\\u0027ctl00\$cp1\$ddCitizenship\\u0027,\\u0027arguments\\u0027\).*?\$get\(""ctl00_cp1_ddCitizenship""\)\);";
            string rgValue = "\"value\":\"(?<value>[^\"]+)\"";
            Match mCitizenship = Regex.Match(content, rgCitizenship);
            if (mCitizenship.Success)
            {
                _listCitizenship.Clear();
                string res = mCitizenship.Value;
                var mValue = Regex.Matches(mCitizenship.Value, rgValue);
                foreach (Match m in mValue)
                {
                    _listCitizenship.Add(new KeyValueItem(m.Groups["value"].ToString(), string.Empty));
                }
            }

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(content);
            HtmlNodeCollection citizenshipNodes = doc.DocumentNode.SelectNodes("//div[@id='ctl00_cp1_ddCitizenship']//li[@class='rcbItem']");

            if (citizenshipNodes != null)
            {
                int count = 0;
                foreach (HtmlNode node in citizenshipNodes)
                {
                    _listCitizenship[count++].Value = RefineString(node.InnerText.Trim());
                }
            }
        }

        private void ParseResidence(string content)
        {
            string rgResidence = @"__doPostBack\(\\u0027ctl00\$cp1\$ddCountryOfResidence\\u0027,\\u0027arguments\\u0027\).*?\$get\(""ctl00_cp1_ddCountryOfResidence""\)\);";
            string rgValue = "\"value\":\"(?<value>[^\"]+)\"";
            Match mResidence = Regex.Match(content, rgResidence);
            if (mResidence.Success)
            {
                _listResidence.Clear();
                string res = mResidence.Value;
                var mValue = Regex.Matches(mResidence.Value, rgValue);
                foreach (Match m in mValue)
                {
                    _listResidence.Add(new KeyValueItem(m.Groups["value"].ToString(), string.Empty));
                }
            }

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(content);
            HtmlNodeCollection residenceNodes = doc.DocumentNode.SelectNodes("//div[@id='ctl00_cp1_ddCountryOfResidence']//li[@class='rcbItem']");

            if (residenceNodes != null)
            {
                int count = 0;
                foreach (HtmlNode node in residenceNodes)
                {
                    _listResidence[count++].Value = RefineString(node.InnerText.Trim());
                }
            }
        }

        private void ParseEmbassy(string content)
        {
            if ("39" == "39")
            {
                _listEmbassy = new List<KeyValueItem>() { new KeyValueItem("129", "Vietnam (Việt Nam) - Hanoj") };
                return;
            }

            string rgEmbassy = @"__doPostBack\(\\u0027ctl00\$cp1\$ddEmbassy\\u0027,\\u0027arguments\\u0027\).*?\$get\(""ctl00_cp1_ddEmbassy""\)\);";
            string rgValue = "\"value\":\"(?<value>[^\"]+)\"";
            Match mEmbassy = Regex.Match(content, rgEmbassy);
            if (mEmbassy.Success)
            {
                _listEmbassy.Clear();
                string res = mEmbassy.Value;
                var mValue = Regex.Matches(mEmbassy.Value, rgValue);
                foreach (Match m in mValue)
                {
                    _listEmbassy.Add(new KeyValueItem(m.Groups["value"].ToString(), string.Empty));
                }
            }

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(content);
            HtmlNodeCollection enbassyNodes = doc.DocumentNode.SelectNodes("//div[@id='ctl00_cp1_ddEmbassy']//li[@class='rcbItem']");

            if (enbassyNodes != null)
            {
                int count = 0;
                foreach (HtmlNode node in enbassyNodes)
                {
                    _listEmbassy[count++].Value = node.InnerText.Trim();
                }
            }
        }
        #endregion

        #region GetRsm1, GetViewState, GetEventValidation
        private string GetRsm1(string content)
        {
            return string.Empty;

            string link = "https://visapoint.eu" + Regex.Match(content, "\"(?<value>/Telerik.Web.UI.WebResource.axd[^\"]*PublicKeyToken[^\"]*)\"").Groups["value"].ToString();
            link = link.Replace("amp;", "");
            string responsejs = http.FetchHttpGet(link, http.LastUrl);

            string tsm = Regex.Match(content, "_TSM_CombinedScripts_=(?<value>[^\"]{10,30})\"").Groups["value"].ToString();

            string rsm1 = EncodeDataString(Regex.Match(responsejs, @"hf\.value \+= '(?<value>[^']*)';").Groups["value"].ToString()) + tsm;
            return rsm1;
        }

        private string GetViewState(string content)
        {
            string result = string.Empty;
            string text = Regex.Match(content, _rgViewState).Groups["value"].ToString();
            string subText = text;
            while (subText.Length > 32766)
            {
                result += EncodeDataString(subText.Substring(0, 32766));
                subText = subText.Substring(32766, subText.Length - 32766);
            }

            result += EncodeDataString(subText);

            return result;
        }

        private string GetEventValidation(string content)
        {
            return EncodeDataString(Regex.Match(content, _rgEventValidation).Groups["value"].ToString());
        }
        #endregion

        private string EncodeDataString(string source)
        {
            return Uri.EscapeDataString(source).Replace("%20", "+").Replace("(", "%28").Replace(")", "%29");
        }

        private string RefineString(string junk)
        {
            // Turn string back to bytes using the original, incorrect encoding.
            byte[] bytes = Encoding.GetEncoding(1252).GetBytes(junk);

            // Use the correct encoding this time to convert back to a string.
            string good = Encoding.UTF8.GetString(bytes);
            return good;
        }

        private static Random RandomProxy = new Random();
        private WebProxy GetRandomProxy()
        {
            if (Setting.ProxyList == null || Setting.ProxyList.Count == 0)
                return null;
            return Setting.ProxyList[RandomProxy.Next(0, Setting.ProxyList.Count)];
        }

        private void WriteLog(string msg)
        {
            new Thread(() =>
                {
                    LogUtils.WriteLogSubmit(string.Format("{0}_{1}", _item.FamilyName, _item.Name), msg);
                }).Start();
        }
        #endregion
    }
}
