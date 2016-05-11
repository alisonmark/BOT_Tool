using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Visa_Appointment_Request;

namespace Auto_VAR
{
    public partial class VarItem
    {
        private string RequestAppointDateTime2(string content, string value)
        {
            string rsm1 = GetRsm1(content);
            string viewState = GetViewState(content);
            string eventValidation = GetEventValidation(content);

            string postData = string.Format("rsm1_TSM={0}"
                                        + "&__EVENTTARGET=ctl00%24cp1%24btnNext"
                                        + "&__EVENTARGUMENT="
                                        + "&__VIEWSTATE={1}"
                                        + "&__VIEWSTATEENCRYPTED="
                                        + "&__EVENTVALIDATION={2}"
                                        + "&ctl00%24ddLocale=English+%28United+Kingdom%29"
                                        + "&ctl00_ddLocale_ClientState="
                                        + "&ctl00%24cp1%24rblDate={3}"
                                        + "&ctl00_cp1_btnPrev_ClientState="
                                        + "&ctl00_cp1_btnNext_ClientState="
                                        , rsm1
                                        , viewState
                                        , eventValidation
                                        , value);
            postData += GetPersonalInfo(content) + GetPassportContact(content);

            return http.FetchHttpPost("https://visapoint.eu/form", http.LastUrl, postData);
        }

        private string GetPersonalInfo(string content)
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

            string postData = string.Format(""
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

            return postData;
        }

        private string GetPassportContact(string content)
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

            string postData = string.Format(""
                                        + "&ctl00%24cp1%24txtPassportNumber={3}"
                                        + "&ctl00_cp1_txtPassportNumber_ClientState={4}"
                                        + "&ctl00%24cp1%24txtEmail={5}"
                                        + "&ctl00_cp1_txtEmail_ClientState={6}"
                                        + "&ctl00%24cp1%24txtPhone={7}"
                                        + "&ctl00_cp1_txtPhone_ClientState={8}"

                                        , rsm1, viewState, eventValidation
                                        , EncodeDataString(passport), passport_ClientState
                                        , EncodeDataString(email), email_ClientState
                                        , EncodeDataString(phone), phone_ClientState);

            return postData;
        }

        private string GetCitizenship(string content)
        {
            int selectedIndex = _listPurposeOfStay.IndexOf(_item.PurposeOfStay);

            string rsm1 = GetRsm1(content);
            string viewState = GetViewState(content);
            string eventValidation = GetEventValidation(content);
            string eventArgument = EncodeDataString(string.Format("{{\"Command\":\"Select\",\"Index\":{0}}}", selectedIndex));
            string clientState = EncodeDataString(string.Format("{{\"logEntries\":[],\"value\":\"{0}\",\"text\":\"{1}\",\"enabled\":true,\"checkedIndices\":[],\"checkedItemsTextOverflows\":false}}", _item.Citizenship.Key, _item.Citizenship.Value));

            string postData = string.Format(""
                                        + "&ctl00%24cp1%24ddCitizenship={4}"
                                        + "&ctl00_cp1_ddCitizenship_ClientState={5}"
                                        , rsm1
                                        , viewState
                                        , eventValidation
                                        , eventArgument
                                        , EncodeDataString(_item.Citizenship.Value)
                                        , clientState);

            return postData;
        }

        private string GetPurposeOfStay(string content)
        {
            KeyValueItem itemEmbassy = _item.Embassy;

            int selectedIndex = _listPurposeOfStay.IndexOf(_item.PurposeOfStay);

            string rsm1 = GetRsm1(content);
            string viewState = GetViewState(content);
            string eventValidation = GetEventValidation(content);
            string eventArgument = EncodeDataString(string.Format("{{\"Command\":\"Select\",\"Index\":{0}}}", selectedIndex));
            string clientState = EncodeDataString(string.Format("{{\"logEntries\":[],\"value\":\"{0}\",\"text\":\"{1}\",\"enabled\":true,\"checkedIndices\":[],\"checkedItemsTextOverflows\":false}}", _item.PurposeOfStay.Key, _item.PurposeOfStay.Value));

            string postData = string.Format(""
                                        + "&ctl00%24cp1%24ddCountryOfResidence={5}"
                                        + "&ctl00_cp1_ddCountryOfResidence_ClientState="
                                        + "&ctl00%24cp1%24ddEmbassy={6}"
                                        + "&ctl00_cp1_ddEmbassy_ClientState="
                                        + "&ctl00%24cp1%24ddVisaType={7}"
                                        + "&ctl00_cp1_ddVisaType_ClientState={8}"
                                        , rsm1
                                        , viewState
                                        , eventValidation
                                        , eventArgument
                                        , EncodeDataString(_item.Citizenship.Value)
                                        , EncodeDataString(_item.Citizenship.Value)
                                        , EncodeDataString(itemEmbassy.Value)
                                        , EncodeDataString(_item.PurposeOfStay.Value)
                                        , clientState);

            return postData;
        }

        private void Step4_AppointDate2()
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
                    WriteLog("Step 4: Không tìm thấy cuộc hẹn nào");
                }

                ShowMessage(this, string.Format("Step 4: Đang chọn ngày hẹn: {0}...", selectedItem.Value));
                WriteLog(string.Format("Step 4: Đang chọn ngày hẹn: {0}...", selectedItem.Value));
                _lastResponse = RequestAppointDateTime2(_lastResponse, selectedItem.Key);
                ShowMessage(this, "Step 4: Đã chọn xong ngày hẹn: " + selectedItem.Value);
                WriteLog("Step 4: Đã chọn xong ngày hẹn: " + selectedItem.Value);

                _item.OpenTime = string.Format("[{0}]{1}", randomIndex, selectedItem.Value);
                Step7_PreviewInfo();
            }
            catch (Exception ex)
            {
                _item.Status = VarState.Error;
                _item.ErrorMsg = ex.Message;
                LogUtils.WriteLog("Step 4: Chọn ngày hẹn gặp lỗi : " + ex.StackTrace, "err");
                ShowMessage(this, "Step 4: Chọn ngày hẹn gặp lỗi : " + ex.Message);
                WriteLog("Step 4: Chọn ngày hẹn gặp lỗi : " + ex.Message);
            }
        }
    }
}
