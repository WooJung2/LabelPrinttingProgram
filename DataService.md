using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Common.DataService
{
    public class DataService
    {
        private OleDbConnection _Conn;
        private Boolean _IsConn = false;

        public DataService()
        {
            _Conn = new OleDbConnection();

            #if DEBUG
            _Conn.ConnectionString = @"provider=microsoft.ace.oledb.12.0;data source=c:\users\hhi\source\repos\labelprinttingsystem\labelprinttingsystem.accdb;persist security info=false;";
            #else
            //_Conn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\LabelPrinttingSystem.accdb;";
            _Conn.ConnectionString = @"provider=microsoft.ace.oledb.12.0;Data Source=.\LabelPrinttingSystem.accdb;";
            #endif
        }

        public Boolean OpenDB()
        {
            Boolean bConn = false;

            try
            {
                _Conn.Open();
                bConn = true;
            }
            catch
            {
                bConn = false;
            }

            _IsConn = bConn;
            return bConn;
        }

        public Boolean CloseDB()
        {
            Boolean bClose = false;

            try
            {
                _Conn.Close();
                bClose = true;
                _IsConn = false;
            }
            catch
            {
                bClose = false;
            }

            return bClose;
        }

        public OleDbConnection GetConObj()
        {
            return _Conn;
        }

        public Boolean IsConn()
        {
            return _IsConn;
        }
    }

    public class DataCommand
    {
        private DataService _ds;
        private Type _typeStr = typeof(string);
        private Type _typeChars = typeof(char[]);
        private Type _typeChar = typeof(char);
        private Type _typeObject = typeof(object);
        private int _errorValue = 0;

        public DataCommand(DataService service)
        {
            _ds = service;
        }

#region Insert

        /// <summary>
        /// Insert ExecuteNonQuery 호출
        /// </summary>
        /// <param name="strTable">대상 테이블명</param>
        /// <param name="dicParameter">입력값</param>
        /// <returns>-1 = SQL 에러</returns>
        public int Insert(string strTable, Dictionary<int, object> dicParameter)
        {
            if (!_ds.IsConn())
                return -1;

            StringBuilder Query = new StringBuilder();
            int nRes = 0;

            Query.Append("INSERT INTO " + strTable + " VALUES (");

            for (int i = 0; i < dicParameter.Count; i++)
            {
                Type type = dicParameter[i].GetType();

                if ((i + 1) > 1)
                {
                    Query.Append(" , ");

                    if(type.Equals(_typeStr) || type.Equals(_typeChars) || type.Equals(_typeChar))
                        Query.Append("'" + dicParameter[i] + "'");
                    else if(type.Equals(typeof(DBNull)))
                        Query.Append("Null");
                    else
                        Query.Append(dicParameter[i]);
                }
                else
                {
                    if (type.Equals(_typeStr) || type.Equals(_typeChars) || type.Equals(_typeChar))
                        Query.Append("'" + dicParameter[i] + "'");
                    else
                        Query.Append(dicParameter[i]);
                }
            }

            Query.Append(")");

            OleDbCommand OleDbCmd = new OleDbCommand(Query.ToStr(), _ds.GetConObj());
            OleDbCmd.CommandType = CommandType.Text;
            nRes = OleDbCmd.ExecuteNonQuery();

            OleDbCmd.Dispose();

            return nRes;
        }

#endregion Insert

#region Update

        /// <summary>
        /// Update ExecuteNonQuery 호출
        /// </summary>
        /// <param name="strTable">대상 테이블명</param>
        /// <param name="dicParameterForWhere">특정 조건</param>
        /// <param name="dicParameterForTarget">Update 대상</param>
        /// <param name="strFlag">Product Update할경우 "P", Vendor 체크박스 Update할 경우 "V"</param>
        /// <param name="strQuery">다중 테이블일 경우 쿼리 필수</param>
        /// <returns>
        /// -1 = SQL 에러
        /// -2 = dicParameterForTarget의 수가 0일 때
        /// </returns>
        public int Update(string strTable, Dictionary<string, object> dicParameterForWhere, Dictionary<string, object> dicParameterForTarget, string strFlag = "P", string strQuery = "")
        {
            if (!_ds.IsConn())
                return -1;

            StringBuilder Query = new StringBuilder();
            int nRes = 0;

            if ("P".Equals(strFlag))
            {
                if (dicParameterForTarget.Count == 0) return -2;

                Query.Append("UPDATE " + strTable + " SET ");

#region Set 조건

                for (int i = 0; i < dicParameterForTarget.Count; i++)
                {
                    Type type = dicParameterForTarget.Keys.ElementAt<string>(i).GetType();

                    if ((i + 1) > 1)
                    {
                        Query.Append(" , " + dicParameterForTarget.Keys.ElementAt<string>(i));

                        if (type.Equals(_typeStr) || type.Equals(_typeChars) || type.Equals(_typeChar))
                            Query.Append(" = '" + dicParameterForTarget.Values.ElementAt<object>(i).ToStr() + "'");
                        else
                            Query.Append(" = " + dicParameterForTarget.Values.ElementAt<object>(i).ToStr());
                    }
                    else
                    {
                        Query.Append(dicParameterForTarget.Keys.ElementAt<string>(i));

                        if(type.Equals(_typeStr) || type.Equals(_typeChars) || type.Equals(_typeChar))
                            Query.Append(" = '" + dicParameterForTarget.Values.ElementAt<object>(i).ToStr() + "'");
                        else
                            Query.Append(" = " + dicParameterForTarget.Values.ElementAt<object>(i).ToStr());
                    }
                }

#endregion Set 조건

#region Where 조건

                for (int i = 0; i < dicParameterForWhere.Count; i++)
                {
                    Type type = dicParameterForWhere.Keys.ElementAt<string>(i).GetType();

                    if ((i + 1) > 1)
                    {
                        Query.Append(" AND " + dicParameterForWhere.Keys.ElementAt<string>(i));

                        if (type.Equals(_typeStr) || type.Equals(_typeChars) || type.Equals(_typeChar))
                            Query.Append(" = '" + dicParameterForWhere.Values.ElementAt<object>(i).ToStr() + "'");
                        else
                            Query.Append(" = " + dicParameterForWhere.Values.ElementAt<object>(i).ToStr());
                    }
                    else
                    {
                        Query.Append(" WHERE " + dicParameterForWhere.Keys.ElementAt<string>(i));

                        if (type.Equals(_typeStr) || type.Equals(_typeChars) || type.Equals(_typeChar))
                            Query.Append(" = '" + dicParameterForWhere.Values.ElementAt<object>(i).ToStr() + "'");
                        else
                            Query.Append(" = " + dicParameterForWhere.Values.ElementAt<object>(i).ToStr());
                    }
                }
            }
            else
            {
                Query.Append(strQuery);
            }

#endregion Where 조건

            OleDbCommand OleDbCmd = new OleDbCommand(Query.ToStr(), _ds.GetConObj());
            OleDbCmd.CommandType = CommandType.Text;
            nRes = OleDbCmd.ExecuteNonQuery();

            OleDbCmd.Dispose();

            return nRes;
        }

#endregion Update

#region Delete

        /// <summary>
        /// Delete ExecuteNonQuery 호출
        /// </summary>
        /// <param name="strTable">대상 테이블명</param>
        /// <param name="dicParameter">Where 조건</param>
        /// <returns>-1 = SQL 에러</returns>
        public int Delete(string strTable, Dictionary<string, object> dicParameter)
        {
            if (!_ds.IsConn())
                return -1;

            StringBuilder Query = new StringBuilder();
            int nRes = 0;

            Query.Append("DELETE FROM " + strTable);

#region Where 조건

            for (int i = 0; i < dicParameter.Count; i++)
            {
                Type type = dicParameter.Keys.ElementAt<string>(i).GetType();

                if ((i + 1) > 1)
                {
                    Query.Append("AND " + dicParameter.Keys.ElementAt<string>(i));

                    if (type.Equals(_typeStr) || type.Equals(_typeChars) || type.Equals(_typeChar))
                        Query.Append(" = '" + dicParameter.Values.ElementAt<object>(i).ToStr() + "'");
                    else
                        Query.Append(" = " + dicParameter.Values.ElementAt<object>(i).ToStr());
                }
                else
                {
                    Query.Append(" WHERE " + dicParameter.Keys.ElementAt<string>(i));

                    if (type.Equals(_typeStr) || type.Equals(_typeChars) || type.Equals(_typeChar))
                        Query.Append(" = '" + dicParameter.Values.ElementAt<object>(i).ToStr() + "'");
                    else
                        Query.Append(" = " + dicParameter.Values.ElementAt<object>(i).ToStr());
                }
            }

#endregion Where 조건

            OleDbCommand OleDbCmd = new OleDbCommand(Query.ToStr(), _ds.GetConObj());
            OleDbCmd.CommandType = CommandType.Text;
            nRes = OleDbCmd.ExecuteNonQuery();

            OleDbCmd.Dispose();

            return nRes;
        }

#endregion Delete

#region Select

        /// <summary>
        /// 조회 DB 호출
        /// </summary>
        /// <param name="strTable">테이블(단일)</param>
        /// <param name="dicParameterForWhere">Where 절</param>
        /// <param name="dicParameterForTarget">Select 절</param>
        /// <param name="strFlag">"S" - 단일 테이블, "M" - 다중 테이블</param>
        /// <param name="strQuery">다중 테이블일 경우 쿼리 필수</param>
        /// <returns></returns>
        public DataSet Select(string strTable, Dictionary<string, object> dicParameterForWhere, Dictionary<int, string> dicParameterForTarget, string strFlag = "S", string strQuery = "")
        {
            DataSet ds = new DataSet();

            if (!_ds.IsConn())
            {
                _errorValue = 0;
                return ds;
            }

            StringBuilder Query = new StringBuilder();

            if ("S".Equals(strFlag))
            {
                Query.Append("SELECT ");

#region 조회 컬럼 선정

                for (int i = 0; i < dicParameterForTarget.Count; i++)
                {
                    if ((i + 1) > 1)
                    {
                        Query.Append(", " + dicParameterForTarget[i]);
                    }
                    else
                    {
                        Query.Append(dicParameterForTarget[i]);
                    }
                }

#endregion 조회 컬럼 선정

                Query.Append(" FROM " + strTable);

#region Where 조건

                for (int i = 0; i < dicParameterForWhere.Count; i++)
                {
                    Type type = dicParameterForWhere.Keys.ElementAt<string>(i).GetType();

                    if ((i + 1) > 1)
                    {
                        Query.Append("AND " + dicParameterForWhere.Keys.ElementAt<string>(i));

                        if (type.Equals(_typeStr) || type.Equals(_typeChars) || type.Equals(_typeChar))
                            Query.Append(" = '" + dicParameterForWhere.Values.ElementAt<object>(i).ToStr() + "'");
                        else
                            Query.Append(" = " + dicParameterForWhere.Values.ElementAt<object>(i).ToStr());
                    }
                    else
                    {
                        Query.Append(" WHERE " + dicParameterForWhere.Keys.ElementAt<string>(i));

                        if (type.Equals(_typeStr) || type.Equals(_typeChars) || type.Equals(_typeChar))
                            Query.Append(" = '" + dicParameterForWhere.Values.ElementAt<object>(i).ToStr() + "'");
                        else
                            Query.Append(" = " + dicParameterForWhere.Values.ElementAt<object>(i).ToStr());
                    }
                }

#endregion Where 조건
            }
            else
            {
                Query.Append(strQuery);
            }

            OleDbDataAdapter OLECmd = new OleDbDataAdapter(Query.ToStr(), _ds.GetConObj());
            OLECmd.Fill(ds);

            if (ds.Tables.Count == 0)
                _errorValue = -1;

            OLECmd.Dispose();            

            return ds;
        }
#endregion Select

        public int ErrorValue()
        {
            return _errorValue;
        }
    }
}
