using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.IO;

namespace ChuongTrinhDiemDanh_DHThaiNguyen.Classes
{
    class GoogleSheet
    {
        public string credentialPath;
        public string spreadSheetID;
        public string appName;

        public GoogleCredential credential;
        public SheetsService service;

        string[] Scopes = { SheetsService.Scope.Spreadsheets };

        public bool Begin(string credentialPath, string spreadSheetID, string appName)
        {
            try
            {
                this.credentialPath = credentialPath;
                this.spreadSheetID = spreadSheetID;
                this.appName = appName;

                using (var stream = new FileStream(credentialPath, FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
                }
                service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = appName,
                });
                return true;
            }
            catch (Exception ex)
            {
                ShowMessageBoxException(ex);
                return false;
            }
        }

        public bool WriteData(string range, ValueRange valueRange)
        {
            try
            {
                var updateRequest = service.Spreadsheets.Values.Update(valueRange, this.spreadSheetID, range);
                updateRequest.ValueInputOption =
                    SpreadsheetsResource.
                    ValuesResource.
                    UpdateRequest.
                    ValueInputOptionEnum.
                    USERENTERED;
                updateRequest.Execute();
                return true;
            }
            catch (Exception ex)
            {
                ShowMessageBoxException(ex);
                return false;
            }
        }

        public ValueRange ReadData(string range)
        {
            try
            {
                SpreadsheetsResource.ValuesResource.GetRequest request =
                                service.Spreadsheets.Values.Get(this.spreadSheetID, range);
                ValueRange response = request.Execute();
                return response;
            }
            catch (Exception ex)
            {
                ShowMessageBoxException(ex);
                return null;
            }
        }

        public IList<ValueRange> ReadBatchData(string[] ranges)
        {
            try
            {
                SpreadsheetsResource.ValuesResource.BatchGetRequest request =
                                service.Spreadsheets.Values.BatchGet(this.spreadSheetID);
                request.Ranges = ranges;
                BatchGetValuesResponse response = request.Execute();
                return response.ValueRanges;
            }
            catch (Exception ex)
            {
                ShowMessageBoxException(ex);
                return null;
            }
        }

        public bool WriteBatchData(IList<ValueRange> valueRanges, string[] ranges)
        {
            try
            {
                string valueInputOption = "";
                BatchUpdateValuesRequest requestBody = new BatchUpdateValuesRequest();
                requestBody.ValueInputOption = valueInputOption;
                requestBody.Data = valueRanges;

                SpreadsheetsResource.ValuesResource.BatchUpdateRequest request =
                    service.Spreadsheets.Values.BatchUpdate(requestBody, this.spreadSheetID);

                BatchUpdateValuesResponse response = request.Execute();
                return true;
            }
            catch (Exception ex)
            {
                ShowMessageBoxException(ex);
                return false;
            }
        }

        private void DebugConsolePrint(object message,
    [System.Runtime.CompilerServices.CallerMemberName] string memberName = "")
        {
#if DEBUG
            System.Diagnostics.Trace.WriteLine($"'{memberName}'" + "\t->\t" + message.ToString());
#endif
        }

        public void ShowMessageBoxException(Exception ex)
        {
            //DebugConsolePrint(ex.Message);
#if DEBUG_SHOW_MESSAGEBOX
            System.Windows.Forms.MessageBox.Show(ex.Message, ex.GetType().Name);
#endif
        }
    }
}
