using MySql.Data.MySqlClient;
using Mysqlx.Crud;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace CostData
{
    class ReceiveData
    {
        public static List<BanquetData> FetchBanquetDataByDate(MySqlConnection connection,string ls_date) {
            var result = new List<BanquetData>();
            string[] parts = ls_date.Split('-');
            int liyear = int.Parse(parts[0]);
            int limonth = int.Parse(parts[1]);
            string ls_banquet_sql = @"WITH base_gc AS (
    SELECT 
        gc.openDateTime,
        gc.checkNum,
        gcs.numGuests,
        gc.firstName,
        gc.checkTotal
    FROM guestcheck gc
    LEFT JOIN guestcheckdetailssumrow gcs ON gcs.checkNum = gc.checkNum
    WHERE LOWER(gc.firstName) = 'banquet'
      AND MONTH(gc.openDateTime) = @limonth and year(gc.openDateTime) = @liyear
),
filtered_details AS (
    SELECT 
        checkNum,
        salesCount,
        salesTotal,
        ROW_NUMBER() OVER (PARTITION BY checkNum ORDER BY itemName) AS rownum
    FROM guestcheckdetails
    WHERE 
        detailType = 1
        AND LOWER(itemName) = 'open food'
        AND UPPER(itemName2) LIKE 'TAO CAN'
        AND itemchname = '自定义食品'
),
detail_sums AS (
    SELECT 
        checkNum,
        SUM(CASE 
            WHEN LOWER(itemName) = 'open food' AND UPPER(itemName2) NOT LIKE 'CHA XIE' AND TIME(openDateTime) < '11:00:00' 
            THEN salesTotal ELSE 0 END) AS Breakfast,
        SUM(CASE 
            WHEN LOWER(itemName) = 'open food' AND UPPER(itemName2) NOT LIKE 'CHA XIE' AND TIME(openDateTime) BETWEEN '11:00:00' AND '15:00:00' 
            THEN salesTotal ELSE 0 END) AS Lunch,
        SUM(CASE 
            WHEN LOWER(itemName) = 'open food' AND UPPER(itemName2) NOT LIKE 'CHA XIE' AND TIME(openDateTime) > '15:00:00' 
            THEN salesTotal ELSE 0 END) AS Dinner,
        SUM(CASE 
            WHEN LOWER(itemName) = 'open food' AND UPPER(itemName2) LIKE 'CHA XIE' 
            THEN salesTotal ELSE 0 END) AS chaxie,
        SUM(CASE WHEN LOWER(itemName) = 'open soda' THEN salesTotal ELSE 0 END) AS bever,
        SUM(CASE WHEN LOWER(itemName) = 'open miscellaneous' THEN salesTotal ELSE 0 END) AS misce,
        SUM(CASE WHEN LOWER(itemName) = 'open beer' THEN salesTotal ELSE 0 END) AS beer,
        SUM(CASE WHEN LOWER(itemName) = 'open equiment rental' THEN salesTotal ELSE 0 END) AS equiment,
        SUM(CASE WHEN LOWER(itemName) = 'open room rental' THEN salesTotal ELSE 0 END) AS room
    FROM guestcheckdetails
    WHERE detailType = 1
    GROUP BY checkNum
),
main_rows AS (
    SELECT 
        bgc.openDateTime,
        bgc.checkNum,
        bgc.numGuests,
        fd.salesCount AS tablescount,
        ROUND(fd.salesTotal / NULLIF(fd.salesCount, 0), 2) AS tablesper,
        ROUND(ds.Breakfast, 2) AS Breakfast,
        ROUND(ds.Lunch, 2) AS Lunch,
        ROUND(ds.Dinner, 2) AS Dinner,
        ROUND(ds.chaxie, 2) AS chaxie,
        ROUND(ds.bever, 2) AS bever,
        ROUND(ds.misce, 2) AS misce,
        ROUND(ds.beer, 2) AS beer,
        ROUND(ds.equiment, 2) AS equiment,
        ROUND(ds.room, 2) AS room,
        bgc.checkTotal
    FROM base_gc bgc
    LEFT JOIN filtered_details fd 
        ON fd.checkNum = bgc.checkNum AND fd.rownum = 1
    LEFT JOIN detail_sums ds 
        ON ds.checkNum = bgc.checkNum
),
extra_rows AS (
    SELECT 
        bgc.openDateTime,
        bgc.checkNum,
        NULL AS numGuests,
        fd.salesCount AS tablescount,
        ROUND(fd.salesTotal / NULLIF(fd.salesCount, 0), 2) AS tablesper,
        NULL AS Breakfast,
        NULL AS Lunch,
        NULL AS Dinner,
        NULL AS chaxie,
        NULL AS bever,
        NULL AS misce,
        NULL AS beer,
        NULL AS equiment,
        NULL AS room,
        NULL AS checkTotal
    FROM base_gc bgc
    JOIN filtered_details fd 
        ON fd.checkNum = bgc.checkNum AND fd.rownum > 1
)
SELECT main_rows.*, (select CONCAT(gcditem.itemname2,gcditem.itemchname ) from guestcheckdetails as gcditem where gcditem.checkNum = main_rows.checkNum and gcditem.detailtype=4 LIMIT 1 )  as paymethod FROM main_rows   
UNION ALL
SELECT extra_rows.*, (select CONCAT(gcditem.itemname2,gcditem.itemchname ) from guestcheckdetails as gcditem where gcditem.checkNum = extra_rows.checkNum and gcditem.detailtype=4 LIMIT 1  ) as paymethod FROM extra_rows 
ORDER BY 
    openDateTime,
    checkNum";
            using (var selectCmd = new MySqlCommand(ls_banquet_sql, connection))
            {
                if (connection.State != ConnectionState.Open)  connection.Open();
                selectCmd.Parameters.AddWithValue("@limonth", limonth);
                selectCmd.Parameters.AddWithValue("@liyear", liyear);
                using (var reader = selectCmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var item = new BanquetData
                        {
                            openDateTime = reader.IsDBNull(reader.GetOrdinal("openDateTime")) ? (DateTime?)null : reader.GetDateTime("openDateTime"),
                            checkNum = reader.IsDBNull(reader.GetOrdinal("checkNum")) ? (long?)null : reader.GetInt64("checkNum"),
                            numGuests = reader.IsDBNull(reader.GetOrdinal("numGuests")) ? (long?)null : reader.GetInt64("numGuests"),
                            tablescount = reader.IsDBNull(reader.GetOrdinal("tablescount")) ? (decimal?)null : reader.GetDecimal("tablescount"),
                            tablesper = reader.IsDBNull(reader.GetOrdinal("tablesper")) ? (decimal?)null : reader.GetDecimal("tablesper"),
                            Breakfast = reader.IsDBNull(reader.GetOrdinal("Breakfast")) ? (decimal?)null : reader.GetDecimal("Breakfast"),
                            Lunch = reader.IsDBNull(reader.GetOrdinal("Lunch")) ? (decimal?)null : reader.GetDecimal("Lunch"),
                            Dinner = reader.IsDBNull(reader.GetOrdinal("Dinner")) ? (decimal?)null : reader.GetDecimal("Dinner"),
                            chaxie = reader.IsDBNull(reader.GetOrdinal("chaxie")) ? (decimal?)null : reader.GetDecimal("chaxie"),
                            bever = reader.IsDBNull(reader.GetOrdinal("bever")) ? (decimal?)null : reader.GetDecimal("bever"),
                            misce = reader.IsDBNull(reader.GetOrdinal("misce")) ? (decimal?)null : reader.GetDecimal("misce"),
                            beer = reader.IsDBNull(reader.GetOrdinal("beer")) ? (decimal?)null : reader.GetDecimal("beer"),
                            equiment = reader.IsDBNull(reader.GetOrdinal("equiment")) ? (decimal?)null : reader.GetDecimal("equiment"),
                            room = reader.IsDBNull(reader.GetOrdinal("room")) ? (decimal?)null : reader.GetDecimal("room"),
                            checkTotal = reader.IsDBNull(reader.GetOrdinal("checkTotal")) ? (decimal?)null : reader.GetDecimal("checkTotal"),
                            paymethod = reader.IsDBNull(reader.GetOrdinal("paymethod")) ? null : reader.GetString("paymethod"),
                        };
                        result.Add(item);
                    }
                }
                
            }
            return result;
        }
        public static List<ChineseFoodData> FetchChineseFoodDataByDate(MySqlConnection connection, string ls_date)
        {
            var result = new List<ChineseFoodData>();
            string[] parts = ls_date.Split('-');
            int liyear = int.Parse(parts[0]);
            int limonth = int.Parse(parts[1]);

            string lssql = @"WITH 
base_details AS(
    SELECT
        checkNum,
        itemName,
        itemName2,
        itemchname,
        salesCount,
        salesTotal,
        openDateTime,
        MENU_ITEM_STRING.MAJORGROUPNAMEMASTER as MAJORGROUP
    FROM guestcheckdetails
    LEFT JOIN MENU_ITEM_STRING
        ON guestcheckdetails.recordID = MENU_ITEM_STRING.MENUITEMID
        AND MENU_ITEM_STRING.POSLANGUAGEID = 3
    WHERE
        guestcheckdetails.detailType = 1
), 
shi_pin_records AS(
    SELECT
        *
    FROM base_details
    WHERE  MAJORGROUP = '食品'
), 
zhuoshu_record AS(
    SELECT checkNum, salesCount, salesTotal
    FROM(
            SELECT checkNum, salesCount, salesTotal,
            ROW_NUMBER() OVER(PARTITION BY checkNum ORDER BY salesTotal DESC) AS rn
        FROM guestcheckdetails
        WHERE itemName = 'Open Food'
    ) AS subquery
    WHERE rn = 1
), 
time_sales AS(
    SELECT
        checkNum,
        ROUND(SUM(CASE WHEN TIME(openDateTime) < '11:00:00' THEN salesTotal ELSE 0 END), 2) AS Breakfast,
        ROUND(SUM(CASE WHEN TIME(openDateTime) BETWEEN '11:00:00' AND '15:00:00' THEN salesTotal ELSE 0 END), 2) AS Lunch,
        ROUND(SUM(CASE WHEN TIME(openDateTime) > '15:00:00' THEN salesTotal ELSE 0 END), 2) AS Dinner
    FROM shi_pin_records
    GROUP BY checkNum
), 
category_sales AS(
    SELECT
        checkNum,
        ROUND(SUM(CASE WHEN MAJORGROUP = '葡萄酒' THEN salesTotal ELSE 0 END), 2) AS wine,
        ROUND(SUM(CASE WHEN MAJORGROUP = '烈性酒'  THEN salesTotal ELSE 0 END), 2) AS liquor,
        ROUND(SUM(CASE WHEN MAJORGROUP = '啤酒'   THEN salesTotal ELSE 0 END), 2) AS beer,
        ROUND(SUM(CASE WHEN MAJORGROUP = '无酒精饮料'  THEN salesTotal ELSE 0 END), 2) AS bever,
        ROUND(SUM(CASE WHEN MAJORGROUP = '杂项' or MAJORGROUP is null THEN salesTotal ELSE 0 END), 2) AS misc
    FROM base_details
    GROUP BY checkNum
) 
SELECT
    gc.openDateTime,
    gc.checkNum,
    gcs.numGuests,
    zs.salesCount AS tablescount,
    zs.salesTotal as tablesper,
    ts.Breakfast,
    ts.Lunch,
    ts.Dinner,
    cs.wine,
    cs.liquor,
    cs.beer,
    cs.bever,
    cs.misc as misce,
    gc.checkTotal AS checktotal
,CONCAT(gcditem.itemname2, gcditem.itemchname) AS paymethod
FROM guestcheck gc
LEFT JOIN guestcheckdetailssumrow gcs ON gcs.checkNum = gc.checkNum
LEFT JOIN zhuoshu_record zs ON zs.checkNum = gc.checkNum
LEFT JOIN time_sales ts ON ts.checkNum = gc.checkNum
LEFT JOIN category_sales cs ON cs.checkNum = gc.checkNum
LEFT JOIN guestcheckdetails gcditem   ON gcditem.checkNum = gc.checkNum   AND gcditem.detailtype = 4
WHERE MONTH(gc.openDateTime) = @limonth AND year(gc.openDateTime) = @liyear and LOWER(gc.firstName) <> 'banquet'
and gc.checkTotal > 0;
";

            using (var selectCmd = new MySqlCommand(lssql, connection))
            {
                if (connection.State != ConnectionState.Open) connection.Open();
                selectCmd.Parameters.AddWithValue("@limonth", limonth);
                selectCmd.Parameters.AddWithValue("@liyear", liyear);
                using (var reader = selectCmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var item = new ChineseFoodData
                        {
                            openDateTime = reader.IsDBNull(reader.GetOrdinal("openDateTime")) ? (DateTime?)null : reader.GetDateTime("openDateTime"),
                            checkNum = reader.IsDBNull(reader.GetOrdinal("checkNum")) ? (long?)null : reader.GetInt64("checkNum"),
                            numGuests = reader.IsDBNull(reader.GetOrdinal("numGuests")) ? (long?)null : reader.GetInt64("numGuests"),
                            tablescount = reader.IsDBNull(reader.GetOrdinal("tablescount")) ? (decimal?)null : reader.GetDecimal("tablescount"),
                            tablesper = reader.IsDBNull(reader.GetOrdinal("tablesper")) ? (decimal?)null : reader.GetDecimal("tablesper"),
                            Breakfast = reader.IsDBNull(reader.GetOrdinal("Breakfast")) ? (decimal?)null : reader.GetDecimal("Breakfast"),
                            Lunch = reader.IsDBNull(reader.GetOrdinal("Lunch")) ? (decimal?)null : reader.GetDecimal("Lunch"),
                            Dinner = reader.IsDBNull(reader.GetOrdinal("Dinner")) ? (decimal?)null : reader.GetDecimal("Dinner"),
                            wine = reader.IsDBNull(reader.GetOrdinal("wine")) ? (decimal?)null : reader.GetDecimal("wine"),
                            bever = reader.IsDBNull(reader.GetOrdinal("bever")) ? (decimal?)null : reader.GetDecimal("bever"),
                            misce = reader.IsDBNull(reader.GetOrdinal("misce")) ? (decimal?)null : reader.GetDecimal("misce"),
                            beer = reader.IsDBNull(reader.GetOrdinal("beer")) ? (decimal?)null : reader.GetDecimal("beer"),
                            liquor = reader.IsDBNull(reader.GetOrdinal("liquor")) ? (decimal?)null : reader.GetDecimal("liquor"),
                            checkTotal = reader.IsDBNull(reader.GetOrdinal("checkTotal")) ? (decimal?)null : reader.GetDecimal("checkTotal"),
                            paymethod = reader.IsDBNull(reader.GetOrdinal("paymethod")) ? null : reader.GetString("paymethod"),
                        };
                        result.Add(item);
                    }
                }
              //  return result;
            }
            return result;
        }
        public static List<Tmpdata> FetchTmpdataByDate(MySqlConnection connection, string datadate)
        //   public static DataTable  FetchTmpdataByDate(MySqlConnection connection, string datadate)
        {
            var result = new List<Tmpdata>();
            var dateStr = datadate;// datadate.ToString("yyyy-MM-dd");

            try
            {
                if (connection.State != ConnectionState.Open)
                    connection.Open();


                try
                {
                    // 1. 删除指定日期的现有数据
                    string deleteSql = @"
                    DELETE FROM tmpdata 
                    WHERE DATE(openDateTime) = @Date OR DATE(FZGOrderPlacedTime) = @Date;";

                    using (var deleteCmd = new MySqlCommand(deleteSql, connection))
                    {
                        deleteCmd.Parameters.AddWithValue("@Date", dateStr);
                        deleteCmd.ExecuteNonQuery();
                    }

                    // 2. 插入新数据
                    string insertSql = @"INSERT INTO tmpdata (
    checkNum, openDateTime, checkTotal, guestnum, firstName, lastName, dayPart, rcsname,
    FZGOrderNumber, FZGThirdPartyOrderNumber, FZGPaymentSerialNumber,
    FZGOrderAmount, FZGTotalPayment, FZGPaymentDiscount, FZGReceivedAmount,
    FZGRefundAmount, FZGTotalDiscountAmount1, FZGTotalDiscountAmount2,
    FZGTotalDiscountAmount3, FZGTotalDiscountAmount4, FZGTotalQuantity,
    FZGCustomerCount, FZGCasher, FZGOrderStatus, FZGInvoiceStatus,
    FZGResouceChannel, FZGPayType, FZGRemark, FZGOrderPlacedTime,
    FZGCheckoutTime, FZGCompletionTime
)
WITH
t1_ranked AS (
    SELECT *, ROW_NUMBER() OVER (PARTITION BY checkTotal ORDER BY openDateTime) AS grp
    FROM vguestcheck
    WHERE DATE(openDateTime) = @Date
		order by openDateTime
),
t2_ranked AS (
    SELECT *, ROW_NUMBER() OVER (PARTITION BY FZGReceivedAmount ORDER BY FZGOrderPlacedTime) AS grp
    FROM fuzhangguiordersum
    WHERE DATE(openDateTime) = @Date
		order by FZGOrderPlacedTime
),
matched AS (
    SELECT
        t1.checkNum, t1.openDateTime, t1.checkTotal,
        t1.guestnum,
        t1.firstName, t1.lastName,
        CASE
            WHEN t1.openDateTime IS NOT NULL THEN
                CASE
                    WHEN HOUR(t1.openDateTime) < 11 THEN '早'
                    WHEN HOUR(t1.openDateTime) < 15 THEN '中'
                    ELSE '晚'
                END
            WHEN t2.FZGOrderPlacedTime IS NOT NULL THEN
                CASE
                    WHEN HOUR(STR_TO_DATE(t2.FZGOrderPlacedTime, '%Y-%m-%d %H:%i:%s')) < 11 THEN '早'
                    WHEN HOUR(STR_TO_DATE(t2.FZGOrderPlacedTime, '%Y-%m-%d %H:%i:%s')) < 15 THEN '中'
                    ELSE '晚'
                END
            ELSE NULL
        END AS dayPart,
        t1.rcsname,
        t2.FZGOrderNumber, t2.FZGThirdPartyOrderNumber, t2.FZGPaymentSerialNumber,
        t2.FZGOrderAmount, t2.FZGTotalPayment, t2.FZGPaymentDiscount, t2.FZGReceivedAmount,
        t2.FZGRefundAmount, t2.FZGTotalDiscountAmount1, t2.FZGTotalDiscountAmount2,
        t2.FZGTotalDiscountAmount3, t2.FZGTotalDiscountAmount4, t2.FZGTotalQuantity,
        t2.FZGCustomerCount, t2.FZGCasher, t2.FZGOrderStatus, t2.FZGInvoiceStatus,
        t2.FZGResouceChannel, t2.FZGPayType, t2.FZGRemark, t2.FZGOrderPlacedTime,
        t2.FZGCheckoutTime, t2.FZGCompletionTime
    FROM t1_ranked t1
    JOIN t2_ranked t2
      ON t1.checkTotal = t2.FZGReceivedAmount AND t1.grp = t2.grp
			order by t1.openDateTime, t2.FZGOrderPlacedTime
),
t1_unmatched AS (
    SELECT *
    FROM t1_ranked
    WHERE checkNum NOT IN (SELECT checkNum FROM matched WHERE checkNum IS NOT NULL)  
),
t2_unmatched AS (
    SELECT *
    FROM t2_ranked
    WHERE FZGOrderNumber NOT IN (SELECT FZGOrderNumber FROM matched WHERE FZGOrderNumber IS NOT NULL)  
)

-- ① 已匹配数据
SELECT * FROM matched

UNION ALL

-- ② guestcheck 未匹配数据
SELECT
    t1.checkNum, t1.openDateTime, t1.checkTotal,
    t1.guestnum,
    t1.firstName, t1.lastName,
    CASE
        WHEN t1.openDateTime IS NOT NULL THEN
            CASE
                WHEN HOUR(t1.openDateTime) < 11 THEN '早'
                WHEN HOUR(t1.openDateTime) < 15 THEN '中'
                ELSE '晚'
            END
        ELSE NULL
    END,
    t1.rcsname,
    NULL, NULL, NULL,
    NULL, NULL, NULL, NULL,
    NULL, NULL, NULL,
    NULL, NULL, NULL, NULL, NULL,
    NULL, NULL, NULL, NULL,
    NULL, NULL, NULL, NULL
FROM t1_unmatched t1

UNION ALL

-- ③ fuzhangguiordersum 未匹配数据
SELECT
    NULL, NULL, NULL, NULL, NULL, NULL,
    CASE
        WHEN t2.FZGOrderPlacedTime IS NOT NULL THEN
            CASE
                WHEN HOUR(STR_TO_DATE(t2.FZGOrderPlacedTime, '%Y-%m-%d %H:%i:%s')) < 11 THEN '早'
                WHEN HOUR(STR_TO_DATE(t2.FZGOrderPlacedTime, '%Y-%m-%d %H:%i:%s')) < 15 THEN '中'
                ELSE '晚'
            END
        ELSE NULL
    END,
    NULL,
    t2.FZGOrderNumber, t2.FZGThirdPartyOrderNumber, t2.FZGPaymentSerialNumber,
    t2.FZGOrderAmount, t2.FZGTotalPayment, t2.FZGPaymentDiscount, t2.FZGReceivedAmount,
    t2.FZGRefundAmount, t2.FZGTotalDiscountAmount1, t2.FZGTotalDiscountAmount2,
    t2.FZGTotalDiscountAmount3, t2.FZGTotalDiscountAmount4, t2.FZGTotalQuantity,
    t2.FZGCustomerCount, t2.FZGCasher, t2.FZGOrderStatus, t2.FZGInvoiceStatus,
    t2.FZGResouceChannel, t2.FZGPayType, t2.FZGRemark, t2.FZGOrderPlacedTime,
    t2.FZGCheckoutTime, t2.FZGCompletionTime
FROM t2_unmatched t2;";

                    using (var insertCmd = new MySqlCommand(insertSql, connection))
                    {
                        insertCmd.Parameters.AddWithValue("@Date", dateStr);
                        insertCmd.ExecuteNonQuery();
                    }

                    // 3. 查询并返回结果
                    string selectSql = @"
                    SELECT * 
                    FROM tmpdata  
                    WHERE DATE(openDateTime) = @Date OR DATE(FZGOrderPlacedTime) = @Date;";

                    using (var selectCmd = new MySqlCommand(selectSql, connection))
                    {
                        selectCmd.Parameters.AddWithValue("@Date", dateStr);

                        //MySqlDataAdapter fzg_adapter = new MySqlDataAdapter(selectCmd);
                        //DataTable fzg_dataTable = new DataTable();
                        //fzg_adapter.Fill(fzg_dataTable);
                        //return fzg_dataTable;

                        using (var reader = selectCmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                 
                                var item = new Tmpdata
                                {
                                    Id = reader.GetInt32("id"),
                                    CheckNum = reader.IsDBNull(reader.GetOrdinal("checkNum")) ? (long?)null : reader.GetInt64("checkNum"),
                                    OpenDateTime = reader.IsDBNull(reader.GetOrdinal("openDateTime")) ? (DateTime?)null : reader.GetDateTime("openDateTime"),
                                    CheckTotal = reader.IsDBNull(reader.GetOrdinal("checkTotal")) ? (decimal?)null : reader.GetDecimal("checkTotal"),
                                    Guestnum = reader.IsDBNull(reader.GetOrdinal("guestnum")) ? (decimal?)null : reader.GetDecimal("guestnum"),
                                    FirstName = reader.IsDBNull(reader.GetOrdinal("firstName")) ? null : reader.GetString("firstName"),
                                    LastName = reader.IsDBNull(reader.GetOrdinal("lastName")) ? null : reader.GetString("lastName"),
                                    Rcsname = reader.IsDBNull(reader.GetOrdinal("rcsname")) ? null : reader.GetString("rcsname"),
                                    Kong1 = reader.IsDBNull(reader.GetOrdinal("kong1")) ? null : reader.GetString("kong1"),
                                    Kong2 = reader.IsDBNull(reader.GetOrdinal("kong2")) ? null : reader.GetString("kong2"),
                                    FZGOrderNumber = reader.IsDBNull(reader.GetOrdinal("FZGOrderNumber")) ? (long?)null : reader.GetInt64("FZGOrderNumber"),
                                    FZGThirdPartyOrderNumber = reader.IsDBNull(reader.GetOrdinal("FZGThirdPartyOrderNumber")) ? null : reader.GetString("FZGThirdPartyOrderNumber"),
                                    FZGPaymentSerialNumber = reader.IsDBNull(reader.GetOrdinal("FZGPaymentSerialNumber")) ? null : reader.GetString("FZGPaymentSerialNumber"),
                                    FZGOrderAmount = reader.IsDBNull(reader.GetOrdinal("FZGOrderAmount")) ? (decimal?)null : reader.GetDecimal("FZGOrderAmount"),
                                    FZGTotalPayment = reader.IsDBNull(reader.GetOrdinal("FZGTotalPayment")) ? (decimal?)null : reader.GetDecimal("FZGTotalPayment"),
                                    FZGPaymentDiscount = reader.IsDBNull(reader.GetOrdinal("FZGPaymentDiscount")) ? (decimal?)null : reader.GetDecimal("FZGPaymentDiscount"),
                                    FZGReceivedAmount = reader.IsDBNull(reader.GetOrdinal("FZGReceivedAmount")) ? (decimal?)null : reader.GetDecimal("FZGReceivedAmount"),
                                    FZGRefundAmount = reader.IsDBNull(reader.GetOrdinal("FZGRefundAmount")) ? (decimal?)null : reader.GetDecimal("FZGRefundAmount"),
                                    FZGTotalDiscountAmount1 = reader.IsDBNull(reader.GetOrdinal("FZGTotalDiscountAmount1")) ? (decimal?)null : reader.GetDecimal("FZGTotalDiscountAmount1"),
                                    FZGTotalDiscountAmount2 = reader.IsDBNull(reader.GetOrdinal("FZGTotalDiscountAmount2")) ? (decimal?)null : reader.GetDecimal("FZGTotalDiscountAmount2"),
                                    FZGTotalDiscountAmount3 = reader.IsDBNull(reader.GetOrdinal("FZGTotalDiscountAmount3")) ? (decimal?)null : reader.GetDecimal("FZGTotalDiscountAmount3"),
                                    FZGTotalDiscountAmount4 = reader.IsDBNull(reader.GetOrdinal("FZGTotalDiscountAmount4")) ? (decimal?)null : reader.GetDecimal("FZGTotalDiscountAmount4"),
                                    FZGTotalQuantity = reader.IsDBNull(reader.GetOrdinal("FZGTotalQuantity")) ? (int?)null : reader.GetInt32("FZGTotalQuantity"),
                                    FZGCustomerCount = reader.IsDBNull(reader.GetOrdinal("FZGCustomerCount")) ? (int?)null : reader.GetInt32("FZGCustomerCount"),
                                    FZGCasher = reader.IsDBNull(reader.GetOrdinal("FZGCasher")) ? null : reader.GetString("FZGCasher"),
                                    FZGOrderStatus = reader.IsDBNull(reader.GetOrdinal("FZGOrderStatus")) ? null : reader.GetString("FZGOrderStatus"),
                                    FZGInvoiceStatus = reader.IsDBNull(reader.GetOrdinal("FZGInvoiceStatus")) ? null : reader.GetString("FZGInvoiceStatus"),
                                    FZGResouceChannel = reader.IsDBNull(reader.GetOrdinal("FZGResouceChannel")) ? null : reader.GetString("FZGResouceChannel"),
                                    FZGPayType = reader.IsDBNull(reader.GetOrdinal("FZGPayType")) ? null : reader.GetString("FZGPayType"),
                                    FZGRemark = reader.IsDBNull(reader.GetOrdinal("FZGRemark")) ? null : reader.GetString("FZGRemark"),
                                    FZGOrderPlacedTime = reader.IsDBNull(reader.GetOrdinal("FZGOrderPlacedTime")) ? null : reader.GetString("FZGOrderPlacedTime"),
                                    FZGCheckoutTime = reader.IsDBNull(reader.GetOrdinal("FZGCheckoutTime")) ? null : reader.GetString("FZGCheckoutTime"),
                                    FZGCompletionTime = reader.IsDBNull(reader.GetOrdinal("FZGCompletionTime")) ? null : reader.GetString("FZGCompletionTime"),
                                    Daypart= reader.IsDBNull(reader.GetOrdinal("daypart")) ? null : reader.GetString("daypart")

                                };
                                result.Add(item);
                            }

                            return result;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"数据库操作出错: {ex.Message}");
                    throw;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine($"数据库操作出错: {ex.Message}");
                throw;
            }
        }


       

    }
    public class BanquetData
    {
        public DateTime? openDateTime { get; set; }
        public long? checkNum { get; set; }
        public long? numGuests { get; set; }
        public decimal? tablescount { get; set; }
        public decimal? tablesper { get; set; }
        public decimal? Breakfast { get; set; }
        public decimal? Lunch { get; set; }
        public decimal? Dinner { get; set; }
        public decimal? chaxie { get; set; }
        public decimal? bever { get; set; }
        public decimal? misce { get; set; }
        public decimal? beer { get; set; }
        public decimal? equiment { get; set; }
        public decimal? room { get; set; }
        public decimal? checkTotal { get; set; }
        public string paymethod { get; set; }
    }

    public class ChineseFoodData
    {
        public DateTime? openDateTime { get; set; }
        public long? checkNum { get; set; }
        //人数
        public long? numGuests { get; set; }
        //桌数
        public decimal? tablescount { get; set; }
        //餐标
        public decimal? tablesper { get; set; }
        public decimal? Breakfast { get; set; }
        public decimal? Lunch { get; set; }
        public decimal? Dinner { get; set; }

        //杂项
        public decimal? misce { get; set; }
        //软饮
        public decimal? bever { get; set; }
        //啤酒
        public decimal? beer { get; set; }
        //红酒
        public decimal? wine { get; set; }
        //烈酒
        public decimal? liquor { get; set; }
        //结账方式
        public string paymethod { get; set; }


        
        public decimal? room { get; set; }
        public decimal? checkTotal { get; set; }
       
    }

    public class Tmpdata
    {
        public int Id { get; set; }
        public long? CheckNum { get; set; }
        public DateTime? OpenDateTime { get; set; }
        public decimal? CheckTotal { get; set; }
        public decimal? Guestnum { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Rcsname { get; set; }
        public string Kong1 { get; set; }
        public string Kong2 { get; set; }
        public long? FZGOrderNumber { get; set; }
        public string FZGThirdPartyOrderNumber { get; set; }
        public string FZGPaymentSerialNumber { get; set; }
        public decimal? FZGOrderAmount { get; set; }
        public decimal? FZGTotalPayment { get; set; }
        public decimal? FZGPaymentDiscount { get; set; }
        public decimal? FZGReceivedAmount { get; set; }
        public decimal? FZGRefundAmount { get; set; }
        public decimal? FZGTotalDiscountAmount1 { get; set; }
        public decimal? FZGTotalDiscountAmount2 { get; set; }
        public decimal? FZGTotalDiscountAmount3 { get; set; }
        public decimal? FZGTotalDiscountAmount4 { get; set; }
        public int? FZGTotalQuantity { get; set; }
        public int? FZGCustomerCount { get; set; }
        public string FZGCasher { get; set; }
        public string FZGOrderStatus { get; set; }
        public string FZGInvoiceStatus { get; set; }
        public string FZGResouceChannel { get; set; }
        public string FZGPayType { get; set; }
        public string FZGRemark { get; set; }
        public string FZGOrderPlacedTime { get; set; }
        public string FZGCheckoutTime { get; set; }
        public string FZGCompletionTime { get; set; }
        public string Daypart { get; set; }
    }
}
