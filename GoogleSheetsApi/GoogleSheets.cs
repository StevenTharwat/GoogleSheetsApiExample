using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System.Collections.Generic;
using System.IO;

namespace GoogleSheetsApi
{
    public class GoogleSheets // install ( Google.Apis.Sheets.V4 )
    {
        #region basics
        static readonly string[] scope = { SheetsService.Scope.Spreadsheets };
        static readonly string aplicationName = "Apies";
        static readonly string sheetsId = "1akxlGHHI2-mr5RxkDxWIVGUfeQNHtGIlp4NL-V3LXCo";
        static SheetsService service;
        GoogleCredential credential;
        #endregion

        static string sheetName = "users";
        static string dataRange = "A1:4000";
        static string DeleteColum = "H";


        public GoogleSheets(string _sheetName)
        {
            sheetName = _sheetName;

            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(scope);
                service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = aplicationName,
                });
            }
            
        }

        public GoogleSheets(string _sheetName,string _DataRange)
        {
            sheetName = _sheetName;
            dataRange = _DataRange;

            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(scope);
                service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = aplicationName,
                });
            }
        }

        public IList<IList<object>> updateData() //update the data
        {
            var range = $"{sheetName}!{dataRange}";
            var request = service.Spreadsheets.Values.Get(sheetsId, range);
            var response = request.Execute();
            Log.write("'Go Online' Get all data");

            statics.DataWithoutFlags = statics.ConvertFromIListToList(response.Values);
            return response.Values;

        }

        public List<int> rowNumberSearchByIDList(string ID) // return (RowNumber) or (empty list) if not found
        {
            List<int> rowNumber = new List<int>();

            var rows = statics.DataWithoutFlags;

            if (rows != null && rows.Count > 0)
            {
                bool ifFound = false;
                string RowNumbersString = string.Empty;

                int counter = 0;
                foreach (var row in rows)
                {

                    counter = counter + 1;
                    string caseId = row[0].ToString(); // 0 -> (id Colum number)
                    if (caseId == $"{ID}") 
                    {
                        rowNumber.Add(counter);
                        RowNumbersString += $"{counter.ToString()} , " ;
                        ifFound = true;
                    }
                }
                if (ifFound) RowNumbersString.Remove(RowNumbersString.Length - 2);
                Log.write($"rowNumberSearchByIDList Search by ID: {ID} , found:{ifFound} , RowNumber:{RowNumbersString}");
            }
            return rowNumber;
        }

        public List<int> SearchByAny(string Any) //return number of (Row) and number of (head)
        {

            var rows = statics.DataWithoutFlags; //data noline
            
            List<int> results = new List<int>();
            results.Add(-1);
            results.Add(-1);

            if (rows != null && rows.Count > 0)
            {
                
                int counter = 0;
                foreach (var row in rows)
                {
                    counter = counter + 1;

                    for(int i = 0; i < row.Count; i++)
                    {
                        if (row[i].ToString() == $"{Any}") 
                        {
                            results.Clear(); 
                            results.Add(counter); 
                            results.Add(i); 
                            Log.write($"SearchByAny search For: {Any} found: True and Return rowNumber:{counter} ColumnNumber:{i}");
                        }
                    }
                }
                if(results[0]==-1 && results[1] == -1)
                {
                    Log.write($"SearchByAny search For: {Any} found: False");
                }
            }
            return results;
        }

        public string OnlineReadRows(int RowNumber , int ColumnNumber) //read a spacific rows from the api //return the (value) or (null) if not found
        {
            var Data = updateData();
            if (Data != null && Data.Count > 0)
            {
                if(ColumnNumber <= Data[0].Count)
                {
                    return Data[ColumnNumber].ToString();
                    Log.write($"OnlineReadRows RowNumber: {RowNumber} ,ColumnNumber: {ColumnNumber} ,and return: {Data[ColumnNumber].ToString()} from the data in the google sheets");

                }
                else return null;
            }
            else
            {
                return null;
            }
        }

        public string OfflineReadRows(List<List<object>> list, int RowNumber , int ColumnNumber) //read from the stored data //return the (value) or (null) if not found
        {
            var Data = list;
            if (Data != null && Data.Count > 0)
            {
                if(ColumnNumber <= Data[0].Count)
                {
                    return Data[ColumnNumber].ToString();
                    Log.write($"OfflineReadRows RowNumber: {RowNumber} ,ColumnNumber: {ColumnNumber} ,and return: {Data[ColumnNumber].ToString()} from the list you give to me");

                }else return null;
            }
            else
            {
                return null;
            }
        }

        public void UpdateRow(string WriteRange, List<object> listOFCellValues) // give range and list of arranged values 
        {
            var range = $"{sheetName}!{WriteRange}";
            var valueRange = new ValueRange();
            var objectList = listOFCellValues;
            valueRange.Values = new List<IList<object>> { objectList };
            var updateRequest = service.Spreadsheets.Values.Update(valueRange, sheetsId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updaterequest = updateRequest.Execute();
            Log.write($"'Go Online' UpdateRow Range: {WriteRange}");

        }

        public void UpdateCell(string writeCell, string value) //put one value in the List
        {
            var range = $"{sheetName}!{writeCell}:{writeCell}";
            var valueRange = new ValueRange();
            var objectList = new List<object>() { value };
            valueRange.Values = new List<IList<object>> { objectList };
            var updateRequest = service.Spreadsheets.Values.Update(valueRange, sheetsId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updaterequest = updateRequest.Execute();
            Log.write($"'Go Online' UpdateCell: {writeCell} by value: {value}");

        }

        public void Delete(List<int> RowNumber)
        {
            if(RowNumber.Count > 0)
            {
                foreach (int number in RowNumber)
                {
                    UpdateCell($"{DeleteColum}{number}", "1"); //replace A by delete colum
                }
            }
            Log.write($"Delete RowNumber: {RowNumber}");
        }

        public void CreateRow(List<object> listOFCellValues) // give list of arranged values 
        {
            var range = $"{sheetName}!{dataRange}";
            var valueRange = new ValueRange();
            var objectList = listOFCellValues; // values in sql
            valueRange.Values = new List<IList<object>> { objectList };
            var appendRequest = service.Spreadsheets.Values.Append(valueRange, sheetsId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendrequest = appendRequest.Execute();
            Log.write("'Go Online' CreateRow");

        }

        public void DeleteByRow(List<List<object>> dataList , int RowNumber ,string FilterStringInLowerCase ,bool boolIfFound) 
        {
            for (int i = 1; i < dataList.Count; i++)
            {

                if (dataList[i][RowNumber].ToString().ToLower().Trim() == FilterStringInLowerCase.ToLower().Trim() && boolIfFound)
                {
                    dataList.RemoveAt(i);
                    i -= 1;
                }
                else if(dataList[i][RowNumber].ToString().ToLower().Trim() != FilterStringInLowerCase.ToLower().Trim() && !boolIfFound)
                {
                    dataList.RemoveAt(i);
                    i -= 1;
                }

            }

            Log.write($"DeleteByRow FilterString: '{FilterStringInLowerCase}' isFound: {boolIfFound} .. in the datalist you give me");

            //updateData();
        }

    }
}
