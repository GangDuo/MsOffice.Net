using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Text;

namespace MsOffice
{
    public abstract class AbstractRunnableSql
    {
        public DataTable Execute(string dataSource, object arg)
        {
            DbProviderFactory factory = DbProviderFactories.GetFactory("System.Data.OleDb");
            using (DbConnection conn = factory.CreateConnection())
            {
                #region
                //
                // Excel用の接続文字列を構築.
                //
                // Providerは、Microsoft.ACE.OLEDB.12.0を使用する事。
                // （JETドライバを利用するとxlsxを読み込む事が出来ない。）
                //
                // Extended Propertiesには、ISAMのバージョン(Excel 12.0)とHDRを指定する。
                // （2003までのxlsの場合はExcel 8.0でISAMバージョンを指定する。）
                // HDRは先頭行をヘッダ情報としてみなすか否かを指定する。
                // 先頭行をヘッダ情報としてみなす場合はYESを、そうでない場合はNOを設定。
                //
                // HDR=NOと指定した場合、カラム名はシステム側で自動的に割り振られる。
                // (F1, F2, F3.....となる)
                //
                #endregion
                DbConnectionStringBuilder builder = factory.CreateConnectionStringBuilder();

                builder["Provider"] = "Microsoft.ACE.OLEDB.12.0";
                builder["Data Source"] = dataSource;
                builder["Extended Properties"] = "Excel 12.0;HDR=YES";

                conn.ConnectionString = builder.ToString();
                conn.Open();

                #region
                //
                // SELECT.
                //
                // 通常のSQLのように発行できる。その際シート指定は
                // [Sheet1$]のように行う。範囲を指定することも出来る。[Sheet1$A1:C7]
                // -------------------------------------------------------------------
                // INSERT
                //
                // こちらも普通のSQLと同じように発行できる。
                // 尚、トランザクションは設定できるが効果は無い。
                // （ロールバックを行ってもデータは戻らない。）
                //
                // また、INSERT,UPDATEはエクセルを開いた状態でも
                // 行う事ができる。
                //
                // データの削除は行う事ができない。（制限）
                //
                #endregion
                using (DbCommand command = conn.CreateCommand())
                {
                    return RunSql(command, arg);
                }
            }
            throw new Exception();
        }

        protected abstract DataTable RunSql(DbCommand command, object arg);
    }
}
