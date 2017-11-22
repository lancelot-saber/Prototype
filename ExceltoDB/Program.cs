using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Configuration;

namespace ExceltoDB
{
    class Program
    {
        private static log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        static void Main(string[] args)
        {
            //フォルダー内のExcelファイルを検索
            string path = ConfigurationManager.AppSettings["constr"];
            string[] files = Directory.GetFiles(path + @"\", "*.xls");
            foreach (string file in files)
            {
                int i; int icount = 0; bool f = false;
                //DBの文字列
                string connString = ConfigurationManager.AppSettings["connString"];
                SqlConnection dbTest = new SqlConnection(connString);

                try
                {
                    dbTest.Open();
                }
                catch (Exception e)
                {
                    Console.WriteLine("DataBaseとの接続が失敗しました。\nネットワーク設定、あるいはパスを確認してください。\n" + e);
                    logger.Fatal("DataBaseとの接続が失敗しました。作業は失敗しました。", e);
                    logger.Info("DataBaseとの接続が失敗しました。"+file+"の作業は失敗しました。");
                }
                //各店舗のMAX車庫IDを求める
                SqlCommand dbTestcmd = new SqlCommand();
                dbTestcmd.Connection = dbTest;
                SqlCommand dbTestcmd2 = new SqlCommand();
                dbTestcmd2.Connection = dbTest;
                dbTestcmd2.CommandText = "SELECT TEMPO_CD,max(SHAKO_ID) AS MAX FROM TM_GARAGE GROUP BY TEMPO_CD";
                SqlDataAdapter maxSHAKOID = new SqlDataAdapter(dbTestcmd2);
                DataSet ds = new DataSet();
                maxSHAKOID.Fill(ds, "SHAKOID");
                DataTable dtID = ds.Tables[0];
                //transcationとcmdの関連
                //Transactionの声明
                SqlTransaction tran = dbTest.BeginTransaction();
                dbTestcmd.Transaction = tran;


                //Excelの文字列
                String strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" + " Data Source=" + "'" + file + "'" + ";Extended Properties ='Excel 12.0;HDR=YES;IMEX=1'";
                OleDbConnection objConn = new OleDbConnection(strConn);
                try
                {
                    //ExcelからDataの読み込み
                    objConn.Open();
                    OleDbCommand objCmdSelect = new OleDbCommand("SELECT * FROM [sheet1$]", objConn);
                    OleDbDataAdapter objAdapter1 = new OleDbDataAdapter();
                    objAdapter1.SelectCommand = objCmdSelect;
                    DataSet myDataSet = new DataSet();
                    objAdapter1.Fill(myDataSet, "XLData");
                    objConn.Close();
                    DataTable dt = myDataSet.Tables[0];
                    logger.Info("Excelファイルの読み込みが出来ました。\n");
                    Console.WriteLine("Excelファイルの読み込みが終わりました。\n");
                    Console.WriteLine();



                    //CREATE_DATEのDataの生成
                    DateTime DT = DateTime.Now;


                    for (i = 0; i < dt.Rows.Count; i++)
                    {
                        String shopCD = dt.Rows[i]["店舗番号"].ToString();
                        String shakoName = dt.Rows[i]["車庫名"].ToString();
                        String shakoNumber = dt.Rows[i]["車庫番号"].ToString();
                        String naigaiFlag = dt.Rows[i]["内外フラグ"].ToString();
                        String deleteFlag = dt.Rows[i]["DELETE_FLG"].ToString();

                        //SHAKO_CDの構造
                        int shakoID = 0;
                        if (dtID.Select("TEMPO_CD='" + shopCD + "'").Length > 0)
                        {
                            DataRow[] matches = dtID.Select("TEMPO_CD='" + shopCD + "'");
                            String strshakoID = matches[0]["MAX"].ToString();
                            shakoID = Convert.ToInt32(strshakoID);
                            shakoID += 1;
                            matches[0]["MAX"] = shakoID;
                        }
                        else
                        {
                            shakoID += 1;
                            DataRow dtIDRow = dtID.NewRow();
                            dtIDRow[0] = shopCD;
                            dtIDRow[1] = shakoID;
                            dtID.Rows.Add(dtIDRow);
                        }

                        //Dataのインサート
                        dbTestcmd.CommandText = "INSERT INTO TM_GARAGE(TEMPO_CD,SHAKO_ID,SHAKO_NAME,SHAKO_NO,NAIGAI_FLG,CREATE_DATE,DELETE_FLG) VALUES ('" + shopCD + "','" + shakoID + "','" + shakoName + "','" + shakoNumber + "','" + naigaiFlag + "','" + DT + "','" + deleteFlag + "');";
                        dbTestcmd.ExecuteNonQuery();
                        icount++;
                    }
                    tran.Commit();
                    Console.WriteLine("データの登録が完了しました。");
                    Console.WriteLine($"登録完了件数：　{icount}件");
                    f = true;

                }

                catch (OleDbException e)
                {
                    Console.WriteLine("Excelファイルからの読み込みが失敗しました。詳しい情報はエラーメッセージに参照してください。\n" + e);
                    logger.Fatal("システムが停止する致命的な障害が発生。作業は失敗しました。", e);
                    logger.Info("Excelファイルからの読み込みが失敗しました。作業は失敗しました。");
                }

                catch (InvalidOperationException e)
                {
                    Console.WriteLine("Excelファイルにアクセスができませんでした。ファイルパスを確認してください。\n" + e);
                    logger.Fatal("システムが停止する致命的な障害が発生。作業は失敗しました。", e);
                    logger.Info("Excelファイルにアクセスができませんでした。作業は失敗しました。");
                }
                catch (SqlException e)
                {
                    Console.WriteLine("登録処理は途中で失敗しました。今回の操作はキャンセルしました。エラーメッセージに参照してください。\nPlease　check　again.\n" + e);
                    tran.Rollback();
                    logger.Fatal("システムが停止する致命的な障害が発生。作業は失敗しました。", e);
                    logger.Info("登録処理は途中で失敗しました。今回の操作はキャンセルしました。作業は失敗しました。");
                }
                catch (Exception e)
                {
                    Console.WriteLine("登録処理は途中で失敗しました。今回の操作はキャンセルしました。\nPlease　check　again.\n エラーメッセージに参照してください。\n" + e);
                    tran.Rollback();
                    logger.Fatal("システムが停止する致命的な障害が発生。作業は失敗しました。", e);
                    logger.Info("登録処理は途中で失敗しました。今回の操作はキャンセルしました。作業は失敗しました。");
                }

                try
                {
                    //ファイルの移動 
                    if (f == true)
                    {
                        string sourcePath = file;
                        string targetPath = ConfigurationManager.AppSettings["targetPath"]+ "車庫" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
                        File.Move(sourcePath, targetPath);
                        Console.WriteLine("Excelファイルはバックアップフォルダに移動しました。");
                        logger.Info("今回" + file + "の作業は成功に実行しました。");
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("Excelファイルはバックアップフォルダの移動は失敗しました。\n" + e);
                    logger.Error("システムが停止するまではいかない障害が発生。作業は成功しましたが、ファイルの移動は失敗しました。", e);
                    logger.Info("作業は成功しましたが、" + file + "のファイル移動は失敗しました。");
                }

                finally
                {
                    dbTest.Close();
                }

            }
        }
    }
}
