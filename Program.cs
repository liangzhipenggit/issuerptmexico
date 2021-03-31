using RM.Common.DotNetData;
using SAP.Middleware.Connector;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using TMSAJobList;
using Aspose.Cells;
namespace MyJobs
{
    class Program 
    {
        public static void Main()
        {
            try
            {

                MaterialIssuePros("CLIENT=888;USER=RFCSCBC;PASSWD=nkY8e9d@#432;LANG=zh;ASHOST=10.118.1.218;SYSNR=00", "Server=10.138.96.99;Database=rm_db_tmsa;Uid=prd_cmp;Pwd=db4CMP2019#");
            }
            catch (Exception e)
            {
                File.WriteAllText("log.txt",e.Message);
            }    
        }
        #region 1 . 收料报表
        public void ReceiptReportIssue(string tmsa_cmp_con)
        {
            SqlHelper helper = new SqlHelper();
            helper.sqltr = tmsa_cmp_con;

            DataSet dtCotent = helper.GetDataSetBySql(@" 
SELECT  
                           CI.ContainerNo, 
                           
                            
                           COUNT(DISTINCT A.PackageNo) AS CTNQTY, 
                           COUNT(DISTINCT B.HQBoxCode) AS REPCTNQTY,
                   CONCAT(SUBSTRING(CONVERT(NVARCHAR(50),CAST(round( COUNT(DISTINCT B.HQBoxCode)*1.0/COUNT(DISTINCT A.PackageNo),4)as decimal(18,4))*100),0,6),'%') 'ReceiptedPercentage',
                    (COUNT(DISTINCT A.PackageNo)-COUNT(DISTINCT B.HQBoxCode)) Balance 
              
                    FROM TMSA_PackingList AS A WITH(NOLOCK)
                         JOIN TMSA_ContainerInfo AS CI WITH(NOLOCK) ON SUBSTRING(A.ContainerNo, 1, 11) = CI.ContainerNo
                                                          AND A.ODF = CI.ODF
                         --LEFT JOIN  (SELECT * FROM VMI_Package_list P where CONVERT(varchar(100), P.CreateDate,20) >='" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + @" 06:30:00.000' --AND CONVERT(varchar(100), P.CreateDate,20) <='" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + @" 16:30:00.000' )  AS B ON A.PackageNo = B.HQBoxCode
                                                         -- AND A.ContainerNo = B.ContainerNo
                                                          --AND A.ODF = B.ODF
                                                          --AND A.Matnr = CI.Material

                   LEFT JOIN  VMI_Package_list B WITH(NOLOCK)  ON A.PackageNo = B.HQBoxCode
                                                            AND A.ContainerNo = B.ContainerNo
                                                            AND   ( A.ODF = CI.ODF
                                                                OR B.ODF = 'Spare parts'
                                                                 )
                                                            AND A.Matnr = B.Material   
                    WHERE CI.DeleteMark = 0   
					 AND  A.ContainerNo IN (SELECT ContainerNo FROM VMI_Package_list AS C WITH(NOLOCK) WHERE  CONVERT(varchar(100), C.CreateDate,20) >='" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + @" 06:30:00.000' AND CONVERT(varchar(100), C.CreateDate,20) <='" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + @" 16:30:00.000')
                     GROUP BY   
                             CI.ContainerNo; 
                          

        SELECT   
                           CI.ContainerNo, 
                           
                           
                            
                           COUNT(DISTINCT A.PackageNo) AS CTNQTY, 
                           COUNT(DISTINCT B.HQBoxCode) AS REPCTNQTY,
                   CONCAT(SUBSTRING(CONVERT(NVARCHAR(50),CAST(round( COUNT(DISTINCT B.HQBoxCode)*1.0/COUNT(DISTINCT A.PackageNo),4)as decimal(18,4))*100),0,6),'%') 'ReceiptedPercentage',
                    (COUNT(DISTINCT A.PackageNo)-COUNT(DISTINCT B.HQBoxCode)) Balance 
              
                    FROM TMSA_PackingList AS A WITH(NOLOCK)
                         JOIN TMSA_ContainerInfo AS CI WITH(NOLOCK) ON SUBSTRING(A.ContainerNo, 1, 11) = CI.ContainerNo
                                                          AND A.ODF = CI.ODF
                         --LEFT JOIN  (SELECT * FROM VMI_Package_list P where CONVERT(varchar(100), P.CreateDate,20) >='" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + @" 16:00:00.000' --AND CONVERT(varchar(100), P.CreateDate,20) <='" + DateTime.Now.ToString("yyyy-MM-dd") + @" 06:00:00.000' )  AS B ON A.PackageNo = B.HQBoxCode
                                                          --AND A.ContainerNo = B.ContainerNo
                                                          --AND A.ODF = B.ODF
                                                         --AND A.Matnr = B.Material
         LEFT JOIN  VMI_Package_list B WITH(NOLOCK)  ON A.PackageNo = B.HQBoxCode
                                                            AND A.ContainerNo = B.ContainerNo
                                                            AND
                                                               ( A.ODF = CI.ODF
                                                                OR B.ODF = 'Spare parts'
                                                                 )
                                                            AND A.Matnr = B.Material   

                    WHERE CI.DeleteMark = 0   AND CI.Status IN('Out','Empty')
					 AND  A.ContainerNo IN (SELECT ContainerNo FROM VMI_Package_list AS C WITH(NOLOCK) WHERE  CONVERT(varchar(100), C.CreateDate,20) >='" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + @" 16:00:00.000' AND CONVERT(varchar(100), C.CreateDate,20) <='" + DateTime.Now.ToString("yyyy-MM-dd") + @" 06:00:00.000')
                     GROUP BY   
                             CI.ContainerNo");
            string mailcontentD = @"";
            string mailcontentN = @"";
            DataTable dtMail = helper.GetTableBySql(@"
            SELECT [Mail_type_Code], 
                   [Mail_type_Name], 
                   [Mail_type_Des], 
                   [Mail_type_To], 
                   [Mail_type_CC], 
                   [Mail_Content], 
                   [Mail_title]
            FROM Base_Maillist WITH(NOLOCK)
            WHERE Mail_type_Code = 'SM04';");





            /*Dayshift*/
            if (dtCotent.Tables.Count > 0)
            {
                foreach (DataRow dr in dtCotent.Tables[0].Rows)
                {

                    if (dr["ReceiptedPercentage"].ToString().Trim() == "100.0%")
                    {
                        mailcontentD += @"<tr  style='background:rgb(132,251,81)'>
					                <td style='border:1px solid black;'>" + dr["ContainerNo"] + @"</td>
					                <td style='border:1px solid black;'>" + dr["CTNQTY"] + @"</td>
					                <td style='border:1px solid black;'>" + dr["REPCTNQTY"] + @"</td>
					                <td style='border:1px solid black;'>" + dr["ReceiptedPercentage"] + @"</td>
					                <td style='border:1px solid black;'>" + dr["Balance"] + @"</td>
				                </tr>";

                    }
                    else if (dr["ReceiptedPercentage"].ToString().Trim() == "0.0%")
                    {
                        mailcontentD += @"<tr  style='background:rgb(254,44,29)'>
					                <td style='border:1px solid black;'>" + dr["ContainerNo"] + @"</td>
					                <td style='border:1px solid black;'>" + dr["CTNQTY"] + @"</td>
					                <td style='border:1px solid black;'>" + dr["REPCTNQTY"] + @"</td>
					                <td style='border:1px solid black;'>" + dr["ReceiptedPercentage"] + @"</td>
					                <td style='border:1px solid black;'>" + dr["Balance"] + @"</td>
				                </tr>";

                    }
                    else
                    {

                        mailcontentD += @"<tr  style='background:rgb(250,254,71)'>
					                <td style='border:1px solid black;'>" + dr["ContainerNo"] + @"</td>
					                <td style='border:1px solid black;'>" + dr["CTNQTY"] + @"</td>
					                <td style='border:1px solid black;'>" + dr["REPCTNQTY"] + @"</td>
					                <td style='border:1px solid black;'>" + dr["ReceiptedPercentage"] + @"</td>
					                <td style='border:1px solid black;'>" + dr["Balance"] + @"</td>
				                </tr>";
                    }

                }
                /*Nightshift*/
                foreach (DataRow dr in dtCotent.Tables[1].Rows)
                {


                    if (dr["ReceiptedPercentage"].ToString().Trim() == "100.0%")
                    {
                        mailcontentN += @"<tr style='background:rgb(132,251,81)'>
					                <td style='border:1px solid black;'>" + dr["ContainerNo"] + @"</td>
 					                <td style='border:1px solid black;'>" + dr["CTNQTY"] + @"</td>
					                <td style='border:1px solid black;'>" + dr["REPCTNQTY"] + @"</td>
					                <td style='border:1px solid black;'>" + dr["ReceiptedPercentage"] + @"</td>
					                <td style='border:1px solid black;'>" + dr["Balance"] + @"</td>
				                </tr>";

                    }

                    else if (dr["ReceiptedPercentage"].ToString().Trim() == "0.0%")
                    {
                        mailcontentN += @"<tr style='background:rgb(254,44,29)'>
					                <td style='border:1px solid black;'>" + dr["ContainerNo"] + @"</td>
 					                <td style='border:1px solid black;'>" + dr["CTNQTY"] + @"</td>
					                <td style='border:1px solid black;'>" + dr["REPCTNQTY"] + @"</td>
					                <td style='border:1px solid black;'>" + dr["ReceiptedPercentage"] + @"</td>
					                <td style='border:1px solid black;'>" + dr["Balance"] + @"</td>
				                </tr>";

                    }
                    else
                    {
                        mailcontentN += @"<tr style='background:rgb(250,254,71)'>
					                <td style='border:1px solid black;'>" + dr["ContainerNo"] + @"</td>
 					                <td style='border:1px solid black;'>" + dr["CTNQTY"] + @"</td>
					                <td style='border:1px solid black;'>" + dr["REPCTNQTY"] + @"</td>
					                <td style='border:1px solid black;'>" + dr["ReceiptedPercentage"] + @"</td>
					                <td style='border:1px solid black;'>" + dr["Balance"] + @"</td>
				                </tr>";

                    }
                }

                //GetSumary
                if (!string.IsNullOrEmpty(mailcontentD))
                {

                    double CNT = 0.00;
                    double RCT = 0.00;
                    double Balance = 0.00;
                    string SumaryRate = string.Empty;
                    GetSumary(dtCotent.Tables[0], out CNT, out RCT, out Balance, out SumaryRate);
                    //<td style='border:1px solid black;'>" + "" + @"</td>

                    if (Balance == 0)
                    {
                        mailcontentD += @"<tr style='background:rgb(132,251,81) '>
					                <td style='border:1px solid black;text-align:right'>Summary" + @"</td>
					                
					                <td style='border:1px solid black;text-align:center'>" + CNT.ToString() + @"</td>
					                <td style='border:1px solid black;text-align:center'>" + RCT.ToString() + @"</td>
					                <td style='border:1px solid black;text-align:center'>" + SumaryRate + @"</td>
					                <td style='border:1px solid black;text-align:center'>" + Balance + @"</td>
				                </tr>";

                    }
                    else
                    {
                        mailcontentD += @"<tr style='background:rgb(250,254,71)'>
					                <td style='border:1px solid black;text-align:right'>Summary" + @"</td>
					                
					                <td style='border:1px solid black;text-align:center'>" + CNT.ToString() + @"</td>
					                <td style='border:1px solid black;text-align:center'>" + RCT.ToString() + @"</td>
					                <td style='border:1px solid black;text-align:center'>" + SumaryRate + @"</td>
					                <td style='border:1px solid black;text-align:center'>" + Balance + @"</td>
				                </tr>";


                    }



                }
                if (!string.IsNullOrEmpty(mailcontentN))
                {

                    double CNT = 0.00;
                    double RCT = 0.00;
                    double Balance = 0.00;
                    string SumaryRate = string.Empty;
                    GetSumary(dtCotent.Tables[1], out CNT, out RCT, out Balance, out SumaryRate);
                    //<td style='border:1px solid black;'>" + "" + @"</td>
                    if (Balance == 0)
                    {
                        mailcontentN += @"<tr style='background:rgb(132,251,81)'>
					                <td style='border:1px solid black;text-align:right'>Summary" + @"</td>
					                
					                <td style='border:1px solid black;text-align:center'>" + CNT.ToString() + @"</td>
					                <td style='border:1px solid black;text-align:center'>" + RCT.ToString() + @"</td>
					                <td style='border:1px solid black;text-align:center'>" + SumaryRate + @"</td>
					                <td style='border:1px solid black;text-align:center'>" + Balance + @"</td>
				                </tr>";

                    }
                    else
                    {
                        mailcontentN += @"<tr style='background:rgb(250,254,71)'>
					                <td style='border:1px solid black;text-align:right'>Summary" + @"</td>
					                
					                <td style='border:1px solid black;text-align:center'>" + CNT.ToString() + @"</td>
					                <td style='border:1px solid black;text-align:center'>" + RCT.ToString() + @"</td>
					                <td style='border:1px solid black;text-align:center'>" + SumaryRate + @"</td>
					                <td style='border:1px solid black;text-align:center'>" + Balance + @"</td>
				                </tr>";


                    }


                }

            }
            if (!string.IsNullOrEmpty(mailcontentD) || !string.IsNullOrEmpty(mailcontentN))
            {
                /*主表内容*/
                string Id = System.Guid.NewGuid().ToString();
                Hashtable ht = new Hashtable();
                ht["ID"] = Id;
                ht["CreateUserId"] = "System";
                ht["CreateUserName"] = "System";
                ht["Flag"] = "0";
                ht["Mail_type_Id"] = "7f9bc21b-9445-4735-b29b-3377b3885001";
                ht["Business_Id"] = "TSCC";
                ht["Mail_type_To"] = dtMail.Rows[0]["Mail_type_To"].ToString();
                ht["Mail_type_CC"] = dtMail.Rows[0]["Mail_type_CC"].ToString();

                ht["Mail_Content"] = @"<html>
                         <head>
                                <style>
                                   table,table tr th, table tr td { border:1px solid black; }
                                   table{ width: 800px; min-height: 25px; line-height: 25px; text-align: center; border-collapse: collapse;} 
                                </style>
                         </head>
                         <body>
                        <p style='font-family:Arial;font-size:small'>
			               Dear leaders and colleagues:
		                </p>
                        <p style='font-family:Arial;font-size:small'>    
                       From Nov 1st, the warehouse began to scan the label to receive raw materials, 
                       the result of scan receipt data as below:
                    
                        </p>" +
                        (!string.IsNullOrEmpty(mailcontentD) ? DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + @"
                        DayShift(06:30~16:30)
		                <table border='1' cellpadding='0' cellspacing='0'>
			                <thead>
				                <tr>
					                <th style='border:1px solid black; font-size:15px;padding:5px'>ContainerNo</th>
					                <th style='border:1px solid black; font-size:15px;padding:5px'>CNT Qty</th>
					                <th style='border:1px solid black; font-size:15px;padding:5px'>Receipt Qty</th>
					                <th style='border:1px solid black; font-size:15px;padding:5px'>Receipt Percentage</th>
					                <th style='border:1px solid black; font-size:15px;padding:5px'>Balance</th>
					                </tr>
			                </thead>
			                <tbody>
				                " + mailcontentD + @"
			                </tbody>
		                </table>" : "") +
                        (!string.IsNullOrEmpty(mailcontentN) ? @"NightShift(16:30~06:30)
                               <table border='1' cellpadding='0' cellspacing='0'>
			                <thead>
				                <tr>
					                <th style='border:1px solid black; font-size:15px;padding:5px'>ContainerNo</th>
					                <th style='border:1px solid black; font-size:15px;padding:5px'>CNT Qty</th>
					                <th style='border:1px solid black; font-size:15px;padding:5px'>Receipt Qty</th>
					                <th style='border:1px solid black; font-size:15px;padding:5px'>Receipt Percentage</th>
					                <th style='border:1px solid black; font-size:15px;padding:5px'>Balance</th>
					                </tr>
			                </thead>
			                <tbody>" + mailcontentN +
                            @"</tbody>
		                </table>
                       " : "") + @"</body>
                         </html>";
                ht["mail_title"] = "Notice about goods receipts data statistics";
                helper.Submit_AddOrEdit("Base_Mail_Data_list", "ID", "", ht);
            }
        }
       
        #endregion
        #region  2 .发料报表
        private static void SaveIssuePros(SqlHelper helper, string sum)
        {
            try
            {
                Hashtable ht = new Hashtable();
                ht.Add("ID", Guid.NewGuid().ToString());
                ht.Add("Day", System.DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"));
                ht.Add("Sum", sum.Split('.')[0]);
                helper.Submit_AddOrEdit("TMSA_IssuePros", "", "", ht);
            }
            catch (Exception e)
            {

            }

        }
        public static void GetSumary(DataTable dt, out double CNT, out double RCT, out double Balance, out string SumaryRate)
        {
            CNT = 0.00;
            RCT = 0.00;
            Balance = 0.00;
            SumaryRate = string.Empty;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                CNT += double.Parse(dt.Rows[i]["CTNQTY"].ToString());
                RCT += double.Parse(dt.Rows[i]["REPCTNQTY"].ToString());
                Balance += double.Parse(dt.Rows[i]["Balance"].ToString());
                var temp = ((RCT / CNT) * 100).ToString();
                SumaryRate = temp.Length > 4 ? temp.Substring(0, 4) + "%" : temp + "%";
            }
        }
        public static void MaterialIssuePros(string tmsa_sap_con, string tmsa_cmp_con)
        {
            SqlHelper helper = new SqlHelper();
            helper.sqltr = tmsa_cmp_con;
            #region 注释
            /*
       SELECT dbo.WorkCenterOpt(c.ARBPL) ARBPL,d.ODF,c.MATNR,m.maktx_en,m.maktx_zh, 
dbo.Get_IssueGoodsQty(c.MATNR,dbo.WorkCenterOpt(c.ARBPL),d.ODF)  Issued
,dbo.Get_ActDemand(c.MATNR,dbo.WorkCenterOpt(c.ARBPL),d.ODF,d.WorkOrder) Demand
FROM ZCMP_MRP_PICKING_4ISSUERPT c WITH(NOLOCK) JOIN(SELECT WorkOrder,ODF FROM TMSA_Plan_demand WITH(NOLOCK) WHERE ODF IN(SELECT  a.ODF   FROM  FWM_ProductHead a WITH(NOLOCK) join  FWM_ProductDetail b WITH(NOLOCK) on a.ID=b.HeadID and a.Status='入库' 
WHERE
a.CreateDate>=CONVERT(datetime,'" + S_time + @"',101)  AND
a.CreateDate<= CONVERT(datetime,'" + E_time + @"',101)


GROUP BY a.ODF )) as d
on c.AUFNR=d.WorkOrder
AND  c.MATNR NOT LIKE 'SVS76-TCL%'
LEFT JOIN Base_Material m ON  m.matnr=c.MATNR


--AND (
--dbo.Get_StrArrayStrOfIndex ( c.MATNR, '-', 1 ) <> '08' 
--AND dbo.Get_StrArrayStrOfIndex ( c.MATNR, '-', 3 ) NOT LIKE 'LP%')
--AND (
-- dbo.Get_StrArrayStrOfIndex ( c.MATNR, '-', 1 ) <> 'T8' 
-- AND dbo.Get_StrArrayStrOfIndex ( c.MATNR, '-', 3 ) NOT LIKE 'LP%')
--AND (
--dbo.Get_StrArrayStrOfIndex ( c.MATNR, '-', 1 ) <> 'M8' 
--AND dbo.Get_StrArrayStrOfIndex ( c.MATNR, '-', 3 ) NOT LIKE 'LP%')
       */
            #endregion

            /*同步料单数据*/
            if (!AsynZCMP_MRP_Data(tmsa_sap_con, helper))
            {
                if (!AsynZCMP_MRP_Data(tmsa_sap_con, helper))
                {
                    return;
                };
            }
            StringBuilder sb = new StringBuilder(string.Format(@" 
SELECT 
   WERKS,AUFNR, KDAUF, ARBPL,MATNR ,SUM(BDMNG)BDMNG,SUM(ENMNG)ENMNG, SUM(CYL)CYL,SUM(POMNG)POMNG 
   INTO #ZCMP_MRP_PICKING_4ISSUERPT  FROM ZCMP_MRP_PICKING_4ISSUERPT GROUP BY  
  WERKS,AUFNR, KDAUF, ARBPL,MATNR  
; SELECT dbo.WorkCenterOpt(c.ARBPL) ARBPL,d.ODF,d.Filt,c.MATNR,m.maktx_en,m.maktx_zh, 
 dbo.Get_IssueGoodsQtyNew(c.MATNR,dbo.WorkCenterOpt(c.ARBPL),d.ODF,d.WorkOrder)  Issued
 ,dbo.Get_ActDemandNew(d.ODF,d.WorkOrder,c.MATNR,d.OffLineQty,dbo.WorkCenterOpt(c.ARBPL)) Demand
 FROM #ZCMP_MRP_PICKING_4ISSUERPT c WITH(NOLOCK) JOIN 
 MesOffLineStatics d
  on c.AUFNR=d.WorkOrder
  AND  c.MATNR NOT LIKE 'SVS76-TCL%'
  LEFT JOIN Base_Material m ON  m.matnr=c.MATNR;
  DROP TABLE #ZCMP_MRP_PICKING_4ISSUERPT;
 "));
            DataTable dtContent = helper.GetTableBySql(sb.ToString());
            dtContent.DefaultView.Sort = "ODF,ARBPL ASC";
            dtContent = dtContent.DefaultView.ToTable();
            List<A_Model> a_list = dtContent.DataTableToList<A_Model>();
            object o = helper.GetObjectBySql("SELECT   DESCRIBE FROM Base_Syscode where ClassCode='GIEX'");
            object oM = helper.GetObjectBySql("SELECT  DESCRIBE FROM Base_Syscode where ClassCode='GIEXM'");
           var M= a_list.Where(c => c.ODF == "DMP9002420" && (
           (c.MATNR.StartsWith("M8")&& c.MATNR.Split('-')[2].StartsWith("LP")) ||
           (c.MATNR.StartsWith("08")&& c.MATNR.Split('-')[2].StartsWith("LP")) || 
           (c.MATNR.StartsWith("T8")&& c.MATNR.Split('-')[2].StartsWith("LP"))
           )).ToList();
            /*排除指定物料  */
            if (oM != null && !string.IsNullOrEmpty(oM.ToString()))
            {
                string[] MATNRarr = oM.ToString().Trim(';').Split(';');
                a_list = a_list.Where(c => !MATNRarr.Contains(c.MATNR.Trim().Split('-')[0]) && !MATNRarr.Contains(c.MATNR)).ToList();
            }
            /*排除指定ODF*/
            if (o != null && !string.IsNullOrEmpty(o.ToString()))
            {
                string[] ODFarr = o.ToString().Trim(';').Split(';');
                a_list = a_list.Where(c => !ODFarr.Contains(c.ODF.Trim())).ToList();
            }
            /*排除模组物料*/
            //var expr = from p in a_list
            //           group p by new { p.ARBPL, p.ODF } into g
            //           where g.FirstOrDefault().ARBPL == "TMSALCM"
            //           select g;

            //List<string> ExLCM = new List<string>();
            //for (int i = 0; i < expr.Count(); i++)
            //{
            //    ExLCM.Add(expr.ToList()[i].Key.ODF);
            //}

            a_list = a_list.Where(c => !(c.Filt=="1" && (
                                  (c.MATNR.Split('-')[0] == "08" && c.MATNR.Split('-')[2].StartsWith("LP")) ||
                                  (c.MATNR.Split('-')[0] == "T8" && c.MATNR.Split('-')[2].StartsWith("LP")) ||
                                  (c.MATNR.Split('-')[0] == "M8" && c.MATNR.Split('-')[2].StartsWith("LP"))))
                           ).Where(c => c.Demand.ToString() != "0.000").Where(f=>!(f.MATNR.StartsWith("40904")||f.MATNR.StartsWith("40903"))).ToList();

            List<CalModel> list = CalculateInfo(a_list, tmsa_cmp_con);
            string mailcontentFA = string.Empty;
            string mailcontentLCM = string.Empty;
            List<CalModel> listFA = list.Where(c => c.ARBPL.Trim() == "TMSAFA").ToList();
            List<CalModel> listLCM = list.Where(c => c.ARBPL.Trim() == "TMSALCM").ToList();
            if (listFA.Count < 1 && listLCM.Count < 1)
            {
                return;
            }
            for (int i = 0; i < listFA.Count; i++)
            {

                if (listFA[i].RATES.Trim() == "100%")
                {
                    mailcontentFA += @"<tr style='background-color:rgb(132,251,81)'>
                   			                <td style='border:1px solid black;text-align:center'>" + listFA[i].ODF + @"</td>
                   			                <td style='border:1px solid black;text-align:center'>" + listFA[i].ARBPL + @"</td>
                                            <td style='border:1px solid black;text-align:center'>" + listFA[i].RATES + @"</td>
                   		           </tr>";
                }

                else if (listFA[i].RATES.Trim() == "0%")
                {
                    mailcontentFA += @"<tr style='background-color:rgb(254,44,29);color:white'>
                   			                <td style='border:1px solid black;text-align:center'>" + listFA[i].ODF + @"</td>
                   			                <td style='border:1px solid black;text-align:center'>" + listFA[i].ARBPL + @"</td>
                                            <td style='border:1px solid black;text-align:center'>" + listFA[i].RATES + @"</td>
                   		           </tr>";
                }
                else
                {
                    mailcontentFA += @"<tr style='background-color:rgb(250,254,71)'>
                   			                <td style='border:1px solid black;text-align:center'>" + listFA[i].ODF + @"</td>
                   			                <td style='border:1px solid black;text-align:center'>" + listFA[i].ARBPL + @"</td>
                                            <td style='border:1px solid black;text-align:center'>" + OptPerct(listFA[i].RATES) + @"</td>
                   		           </tr>";
                }
            }
            /*遍历生成LCM 的数据*/
            for (int i = 0; i < listLCM.Count; i++)
            {

                if (listLCM[i].RATES.Trim() == "100%")
                {
                    mailcontentLCM += @"<tr style='background-color:rgb(132,251,81)'>
                   			                <td style='border:1px solid black;text-align:center'>" + listLCM[i].ODF + @"</td>
                   			                <td style='border:1px solid black;text-align:center'>" + listLCM[i].ARBPL + @"</td>
                                            <td style='border:1px solid black;text-align:center'>" + listLCM[i].RATES + @"</td>
                   		           </tr>";
                }

                else if (listLCM[i].RATES.Trim() == "0%")
                {
                    mailcontentLCM += @"<tr style='background-color:rgb(254,44,29);color:white'>
                   			                <td style='border:1px solid black;text-align:center'>" + listLCM[i].ODF + @"</td>
                   			                <td style='border:1px solid black;text-align:center'>" + listLCM[i].ARBPL + @"</td>
                                            <td style='border:1px solid black;text-align:center'>" + listLCM[i].RATES + @"</td>
                   		           </tr>";
                }
                else
                {
                    mailcontentLCM += @"<tr style='background-color:rgb(250,254,71)'>
                   			                <td style='border:1px solid black;text-align:center'>" + listLCM[i].ODF + @"</td>
                   			                <td style='border:1px solid black;text-align:center'>" + listLCM[i].ARBPL + @"</td>
                                            <td style='border:1px solid black;text-align:center'>" + OptPerct(listLCM[i].RATES) + @"</td>
                   		           </tr>";
                }
            }
            /*遍历生成LCM 的数据End*/
            if (!string.IsNullOrEmpty(mailcontentFA))
            {
                var tFA = listFA.Sum(c => Convert.ToDecimal(c.RATES.ToString().Trim('%')));
                var sumFA = listFA.Sum(c => Convert.ToDecimal(c.RATES.ToString().Trim('%'))) / listFA.Count;
                SaveIssuePros(helper, sumFA.ToString());
                if (OptPerct(sumFA.ToString() + "%") == "100%")
                {
                    mailcontentFA += @"<tr style='background-color:rgb(132,251,81)'>
                   			                <td style='border:1px solid black;text-align:right;'colspan=2 >summary </td>
                   			                 
                                            <td style='border:1px solid black;text-align:center'>" + OptPerct(sumFA.ToString()) + @"</td>
                   		           </tr>";
                }
                else

                {
                    mailcontentFA += @"<tr style='background-color:rgb(250,254,71)'>
                   			                <td style='border:1px solid black;text-align:right;'colspan=2 >summary </td>
                   			                 
                                            <td style='border:1px solid black;text-align:center'>" + OptPerct(sumFA.ToString()) + @"</td>
                   		           </tr>";
                }
            }
            /*颜色标记LCM 数据*/
            if (!string.IsNullOrEmpty(mailcontentLCM))
            {
                var tLCM = listLCM.Sum(c => Convert.ToDecimal(c.RATES.ToString().Trim('%')));
                var sumLCM = listLCM.Sum(c => Convert.ToDecimal(c.RATES.ToString().Trim('%'))) / listLCM.Count;
                SaveIssuePros(helper, sumLCM.ToString());
                if (OptPerct(sumLCM.ToString() + "%") == "100%")
                {
                    mailcontentLCM += @"<tr style='background-color:rgb(132,251,81)'>
                   			                <td style='border:1px solid black;text-align:right;'colspan=2 >summary </td>
                   			                 
                                            <td style='border:1px solid black;text-align:center'>" + OptPerct(sumLCM.ToString() + "%") + @"</td>
                   		           </tr>";
                }
                else

                {
                    mailcontentLCM += @"<tr style='background-color:rgb(250,254,71)'>
                   			                <td style='border:1px solid black;text-align:right;'colspan=2 >summary </td>
                   			                 
                                            <td style='border:1px solid black;text-align:center'>" + OptPerct(sumLCM.ToString() + "%") + @"</td>
                   		           </tr>";
                }
            }

            string allContent = @"
                         <html>
                         <head>
                                <style>
                                   table,table tr th, table tr td { border:1px solid black; }
                                   table{ width: 800px; min-height: 25px; line-height: 25px; text-align: center; border-collapse: collapse;} 
                                </style>
                         </head>
                         <body>
                        <p style='font-family:Arial;font-size:small'>
			               Dear leaders and colleagues:
		                </p>
                        <p style='font-family:Arial;font-size:small'>    
                         This is the statistics data about  material issuing (ODFs which are on producing).please check and confirm.</p>
                          <p style='font-family:Arial;font-size:small' >
                          TMSAFA: " +
                         DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + @"</br>
                          </p>
		                <table border='1' cellpadding='0' cellspacing='0'>
			                <thead>
				                <tr>
					                <th style='border:1px solid black; font-size:15px;padding:5px'>ODF</th>
					                <th style='border:1px solid black; font-size:15px;padding:5px'>Work Center</th>
                                    <th style='border:1px solid black; font-size:15px;padding:5px'>Issued Percentage</th>
			 
 					                </tr>
			                </thead>
			                <tbody>
				                " + mailcontentFA + @"
			                </tbody>
		                </table>
 <p style='font-family:Arial;font-size:small'> 
     TMSALCM: " +
                         DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + @" </p>
   <table border='1' cellpadding='0' cellspacing='0'>
			                <thead>
				                <tr>
					                <th style='border:1px solid black; font-size:15px;padding:5px'>ODF</th>
					                <th style='border:1px solid black; font-size:15px;padding:5px'>Work Center</th>
                                    <th style='border:1px solid black; font-size:15px;padding:5px'>Issued Percentage</th>
			 
 					                </tr>
			                </thead>
			                <tbody>
				                " + mailcontentLCM + @"
			                </tbody>
		                </table>
                       </body>
                         </html>";

            DataTable dtMail = helper.GetTableBySql(@"
            SELECT [Mail_type_Code], 
                   [Mail_type_Name], 
                   [Mail_type_Des], 
                   [Mail_type_To], 
                   [Mail_type_CC], 
                   [Mail_Content], 
                   [Mail_title]
            FROM Base_Maillist
            WHERE Mail_type_Code = 'SM06';");

           SenAttachment(a_list, dtMail.Rows[0]["Mail_type_To"].ToString(), dtMail.Rows[0]["Mail_type_CC"].ToString(), allContent, dtMail.Rows[0]["Mail_title"].ToString(), helper);


        }
        /// <summary>
        /// Spare parts
        /// </summary>
        /// <param name="tmsa_cmp_con"></param>
        /// 
        public static List<CalModel> CalculateInfo(List<A_Model> alist, string tmsa_cmp_con)
        {
            SqlHelper helper = new SqlHelper();
            helper.sqltr = tmsa_cmp_con;
            string ODFWC = string.Empty;
            int p = 0;
            double rat = 0.00;
            double Flag = 0;
            List<CalModel> list = new List<CalModel>();
            for (int i = 0; i < alist.Count; i++)
            {
                if (ODFWC != alist[i].ODF.ToString().Trim() + alist[i].ARBPL.ToString().Trim())
                {
                    if (Flag == 1)
                    {

                        list.Add(new CalModel
                        {
                            ARBPL = alist[i - 1].ARBPL.ToString().Trim(),
                            ODF = alist[i - 1].ODF.ToString().Trim(),
                            RATES = (Math.Round((rat / p), 5) * 100).ToString() + "%"
                        });
                        rat = 0.00;
                        p = 0;
                    }
                    //var R = alist[i];
                    ODFWC = alist[i].ODF.ToString().Trim() + alist[i].ARBPL.ToString().Trim();
                    double Rt = Math.Round(double.Parse(alist[i].Issued.ToString()) / double.Parse(alist[i].Demand.ToString()), 5);
                    if (Rt.ToString() == "NaN")
                    {
                        Rt = 0.00;
                    }
                    if (Rt > 1)
                    {
                        Rt = 1.00;
                    }
                    rat += Rt;
                    p++;
                }
                else
                {
                    ODFWC = alist[i].ODF.ToString().Trim() + alist[i].ARBPL.ToString().Trim();
                    double Rt = Math.Round(double.Parse(alist[i].Issued.ToString()), 5) / Math.Round(double.Parse(alist[i].Demand.ToString()), 5);
                    if (Rt.ToString() == "NaN")
                    {
                        Rt = 0.00;
                    }
                    if (Rt > 1)
                    {
                        Rt = 1.00;
                    }
                    rat += Rt;
                    p++;
                }
                Flag = 1;
            }
            return list;
        }
        public static string OptPerct(string arg)
        {
            int len = arg.Trim().Length;
            string res = arg;
            if (len > 5)
            {
                res = arg.Substring(0, 5) + "%";
            }
            return res;
        }
        public static bool SenAttachment(List<A_Model> alist, string To, string CC, string Content, string Title, SqlHelper helper)
        {
            /*主表内容*/
            string Id = System.Guid.NewGuid().ToString();
            Hashtable ht = new Hashtable();
            ht["ID"] = Id;
            ht["CreateUserId"] = "System";
            ht["CreateUserName"] = "System";
            ht["Flag"] = "0";
            ht["Mail_type_Id"] = "7f9bc21b-9445-4735-b29b-3377b3885001";
            ht["Business_Id"] = "TSCC";
            ht["Mail_type_To"] = To;
            ht["Mail_type_CC"] = CC;
            ht["Mail_Content"] = Content;
            ht["mail_title"] = Title;
            ht["Attachment_Id"] = Id;
            /*主表内容*/
            string FileName;

            MemoryStream ms = SaveFile(alist, out FileName);
            byte[] bt = ms.ToArray();

            IList<System.Data.SqlClient.SqlParameter> parAtt = new List<System.Data.SqlClient.SqlParameter>();
            parAtt.Add(new System.Data.SqlClient.SqlParameter("@ID", Id));
            parAtt.Add(new System.Data.SqlClient.SqlParameter("@AttachmentName", FileName));
            parAtt.Add(new System.Data.SqlClient.SqlParameter("@AttachmentType", "application/vnd.ms-excel"));

            parAtt.Add(new System.Data.SqlClient.SqlParameter("@FileContent", bt));
            StringBuilder strparAtt = new StringBuilder(@" INSERT INTO dbo.Base_Mail_Data_Attachment
                                                       (Id
                                                       ,AttachmentName
                                                       ,AttachmentType 
                                                       ,FileContent,AttachmentUrl)
                                                 VALUES
                                                       ( @ID
                                                        ,@AttachmentName
                                                        ,@AttachmentType 
                                                        ,@FileContent,'N') ");

            return helper.ExecuteBySql(strparAtt, parAtt.ToArray()) > 0 && helper.Submit_AddOrEdit("Base_Mail_Data_list", "ID", "", ht);
        }
        public static bool AsynZCMP_MRP_Data(String sap_con, SqlHelper helper)
        {
            try
            {
                RfcDestination SapDest = null;
                SapDest = SapDesInit.GetDest(sap_con);

                RfcRepository RfcRep = SapDest.Repository;
                IRfcFunction ZCMP_MRP = RfcRep.CreateFunction("ZCMP_MRP");
                SqlHelper helperMes = new SqlHelper();
                helperMes.sqltr = "Server=10.138.96.109;Database=IDOtherData;Uid=prd_smes;Pwd=db4SMES2019#";

                DataTable tblOffLine = helperMes.GetTableBySql(string.Format(@"SELECT DISTINCT  WO.WorkOrderNo
INTO #TEMP_WO
FROM TVInfoBak M WITH(NOLOCK)
LEFT JOIN WorkOrder WO WITH(NOLOCK) ON M.PWorkOrderNo=WO.PWorkOrderNo
WHERE M.EndTime BETWEEN     dateadd(day,(-1),getdate()) and GETDATE() and WO.WorkOrderNo NOT LIKE '0000300%'
GROUP BY WO.WorkOrderNo; 

SELECT WO.WorkOrderNo,WO.ODF,COUNT(M.TVSN) Qty,WO.TMSAODF 
FROM TVInfoBak M WITH(NOLOCK)
LEFT JOIN WorkOrder WO WITH(NOLOCK) ON M.PWorkOrderNo=WO.PWorkOrderNo
WHERE EXISTS(SELECT 1 FROM #TEMP_WO TWO WHERE WO.WorkOrderNo= TWO.WorkOrderNo)
GROUP BY WO.WorkOrderNo,WO.ODF,WO.TMSAODF 
DROP TABLE #TEMP_WO"));

                //获取工厂代码
                DataTable dtFac = helper.GetTableBySql("SELECT Organization_Code FROM [dbo].[Base_Organization] WHERE ParentId IN(SELECT Organization_ID FROM [dbo].[Base_Organization] WHERE Organization_Code = '0988')");
                IRfcTable IT_WEKRS = ZCMP_MRP.GetTable("IT_WEKRS", true);
                IRfcTable MrpItems = null;
                foreach (DataRow drFac in dtFac.Rows)
                {
                    IRfcStructure ZSPP_CMP_WERKS = RfcRep.GetStructureMetadata("ZSPP_CMP_WERKS").CreateStructure();
                    ZSPP_CMP_WERKS.SetValue("WERKS", drFac[0].ToString());
                    IT_WEKRS.Insert(ZSPP_CMP_WERKS);
                }
                /*根据SO 查询是否需要排除物料*/
                tblOffLine.Columns.Add("WRKCT");


                List<FilterModel> fmoel = new List<FilterModel>();

                DataView dataView = tblOffLine.DefaultView;

                DataTable tblDist = dataView.ToTable(true, "TMSAODF");

                if (tblDist.Rows.Count > 0)
                {
                    for (int i = 0; i < tblDist.Rows.Count; i++)
                    {
                        IRfcTable IT_ODF = ZCMP_MRP.GetTable("IT_KDAUF", true);
                        IRfcStructure ZSPP_CMP_ODF = RfcRep.GetStructureMetadata("ZSPP_CMP_KDAUF").CreateStructure();
                        ZSPP_CMP_ODF.SetValue("KDAUF", tblDist.Rows[i]["TMSAODF"]);
                        IT_ODF.Insert(ZSPP_CMP_ODF, i);
                    }
                    ZCMP_MRP.SetValue("I_R_OPTION", "1");
                    ZCMP_MRP.Invoke(SapDest);
                    MrpItems = ZCMP_MRP.GetTable("ET_MRP");



                    for (int i = 0; i < MrpItems.Count; i++)
                    {
                        if (!fmoel.Contains(new FilterModel { ODF = MrpItems[i].GetString("KDAUF"), WORKCENTER = MrpItems[i].GetString("ARBPL").Contains("TMSAFA") ? "TMSAFA" : "TMSALCM" }))
                        {
                            fmoel.Add(new FilterModel { ODF = MrpItems[i].GetString("KDAUF"), WORKCENTER = MrpItems[i].GetString("ARBPL").Contains("TMSAFA") ? "TMSAFA" : "TMSALCM" });
                        } 
                           
                    }
                }
                  /*同步捡料清单*/
                  
                  StringBuilder sb_ZCMP_MRP_PICKING_4ISSUERPT = new StringBuilder("Delete FROM ZCMP_MRP_PICKING_4ISSUERPT");
                if (helper.ExecuteBySql(sb_ZCMP_MRP_PICKING_4ISSUERPT, null) < 0)
                {
                    return false;
                }
                if (tblOffLine.Rows.Count > 0)
                {

                    IRfcTable IT_AUFNR = ZCMP_MRP.GetTable("IT_AUFNR", true);
                    for (int i = 0; i < tblOffLine.Rows.Count; i++)
                    {
                        IRfcStructure ZSPP_CMP_AUFNR = RfcRep.GetStructureMetadata("ZSPP_CMP_AUFNR").CreateStructure();
                        ZSPP_CMP_AUFNR.SetValue("AUFNR", tblOffLine.Rows[i]["WorkOrderNo"]);
                        IT_AUFNR.Insert(ZSPP_CMP_AUFNR, i);
                    }
                }
                else
                {
                    return false;
                }

                SynPICK_LIST(sap_con, helper, tblOffLine);

               
              
                ZCMP_MRP.SetValue("I_R_OPTION", "1");
                ZCMP_MRP.Invoke(SapDest);
                MrpItems = ZCMP_MRP.GetTable("ET_MRP");
                int index = 0;


                DataTable dt = new DataTable();
                dt.TableName = "ZCMP_MRP_PICKING_4ISSUERPT";
                dt.Columns.Add(new DataColumn("WERKS"));
                dt.Columns.Add(new DataColumn("MATNR"));
                dt.Columns.Add(new DataColumn("AUFNR"));
                dt.Columns.Add(new DataColumn("KWMENG"));
                dt.Columns.Add(new DataColumn("KDAUF"));
                dt.Columns.Add(new DataColumn("KDPOS"));
                dt.Columns.Add(new DataColumn("ARBPL"));
                dt.Columns.Add(new DataColumn("BDMNG"));
                dt.Columns.Add(new DataColumn("ENMNG"));
                dt.Columns.Add(new DataColumn("CYL"));
                dt.Columns.Add(new DataColumn("POTX1"));
                dt.Columns.Add(new DataColumn("POMNG"));
                dt.Columns.Add(new DataColumn("PLNBEZ"));
                dt.Columns.Add(new DataColumn("MAABC"));
                dt.Columns.Add(new DataColumn("BAUGR"));
                dt.Columns.Add(new DataColumn("MAKTX"));
                dt.Columns.Add(new DataColumn("PROJN"));
                dt.Columns.Add(new DataColumn("SST3DT"));
                dt.Columns.Add(new DataColumn("ABLAD"));
                dt.Columns.Add(new DataColumn("CGY"));
                dt.Columns.Add(new DataColumn("LGORT"));
                dt.Columns.Add(new DataColumn("RSNUM"));
                dt.Columns.Add(new DataColumn("RSPOS"));
                dt.Columns.Add(new DataColumn("FEVOR"));
                dt.Columns.Add(new DataColumn("CreateDate"));
                dt.Columns.Add(new DataColumn("BDTER"));

                //MrpItems.Where(c => c.GetString("ARBPL").Contains("TMSALCM"));



                if (MrpItems.Count > 0)
                {
                    foreach (var MrpItem in MrpItems)
                    {
                        DataRow dr = dt.NewRow();
                        dr["WERKS"] = MrpItem.GetString("WERKS");
                        dr["MATNR"] = MrpItem.GetString("MATNR");
                        dr["AUFNR"] = MrpItem.GetString("AUFNR");
                        dr["KWMENG"] = MrpItem.GetString("KWMENG");
                        dr["KDAUF"] = MrpItem.GetString("KDAUF");
                        dr["KDPOS"] = MrpItem.GetString("KDPOS");
                        dr["ARBPL"] = MrpItem.GetString("ARBPL").Contains("TMSAFA") ? "TMSAFA" : "TMSALCM";
                        dr["BDMNG"] = MrpItem.GetString("BDMNG");
                        dr["ENMNG"] = MrpItem.GetString("ENMNG");
                        dr["CYL"] = MrpItem.GetString("CYL");
                        dr["POTX1"] = MrpItem.GetString("POTX1");
                        dr["POMNG"] = MrpItem.GetString("POMNG");
                        dr["PLNBEZ"] = MrpItem.GetString("PLNBEZ");
                        dr["MAABC"] = MrpItem.GetString("MAABC");
                        dr["BAUGR"] = MrpItem.GetString("BAUGR");
                        dr["MAKTX"] = MrpItem.GetString("MAKTX");
                        dr["PROJN"] = MrpItem.GetString("PROJN");
                        dr["SST3DT"] = MrpItem.GetString("SST3DT");
                        dr["ABLAD"] = MrpItem.GetString("ABLAD");
                        dr["CGY"] = MrpItem.GetString("CGY");
                        dr["LGORT"] = MrpItem.GetString("LGORT");
                        dr["RSNUM"] = MrpItem.GetString("RSNUM");
                        dr["RSPOS"] = MrpItem.GetString("RSPOS");
                        dr["FEVOR"] = MrpItem.GetString("FEVOR");
                        dr["CreateDate"] = DateTime.Now.ToLocalTime();
                        try
                        {
                            dr["BDTER"] = DateTime.Parse(MrpItem.GetString("BDTER"));
                        }
                        catch
                        {
                            dr["BDTER"] = DBNull.Value;
                        }
                        dt.Rows.Add(dr);
                    }
                }
               var MM= helper.MsSqlBulkCopyData(dt);
                helper.ExecuteBySql(new StringBuilder("EXEC [dbo].[pro_UpZCMP_GI_LIST_RPT]"), null);

                /*获取MES下线数据*/
                List<StringBuilder> lsb = new List<StringBuilder>();
                List<object> lparam = new List<object>();
                var trt = fmoel.Where(c => c.ODF == "0187004224").ToList();
                if (tblOffLine.Rows.Count > 0)
                {
                    lsb.Add(new StringBuilder(" DELETE FROM MesOffLineStatics"));
                    lparam.Add(null);

                    for (int i = 0; i < tblOffLine.Rows.Count; i++)
                    {
                        //                        lsb.Add(new StringBuilder(string.Format(@"INSERT INTO [dbo].[MesOffLineStatics] ([ODF], [WorkOrder], [OffLineQty],[SO]  )
                        //VALUES
                        //	(   N'{0}', N'{1}',{2} );", tblOffLine.Rows[i]["ODF"], tblOffLine.Rows[i]["WorkOrderNo"], tblOffLine.Rows[i]["Qty"], tblOffLine.Rows[i]["TMSAODF"])));
                        //                        lparam.Add(null);
             

                        if (fmoel.Where(c => (c.WORKCENTER == "TMSAFA" && c.ODF == tblOffLine.Rows[i]["TMSAODF"].ToString()
                         .Trim()) || (c.WORKCENTER == "TMSALCM" && c.ODF == tblOffLine.Rows[i]["TMSAODF"].ToString().Trim())).Count() > 1)
                        {
                            lsb.Add(new StringBuilder(string.Format(@"INSERT INTO [dbo].[MesOffLineStatics] ([ODF], [WorkOrder], [OffLineQty],[SO],[Filt]  )
VALUES
	(   N'{0}', N'{1}',{2},'{3}',{4} );", tblOffLine.Rows[i]["ODF"], tblOffLine.Rows[i]["WorkOrderNo"], tblOffLine.Rows[i]["Qty"], tblOffLine.Rows[i]["TMSAODF"],1)));
                            lparam.Add(null);

                        }
                        else
                        {
                            lsb.Add(new StringBuilder(string.Format(@"INSERT INTO [dbo].[MesOffLineStatics] ([ODF], [WorkOrder], [OffLineQty],[SO],[Filt]  )
VALUES
	(   N'{0}', N'{1}',{2},'{3}',{4} );", tblOffLine.Rows[i]["ODF"], tblOffLine.Rows[i]["WorkOrderNo"], tblOffLine.Rows[i]["Qty"], tblOffLine.Rows[i]["TMSAODF"],0)));
                            lparam.Add(null);
                        }
                        
                    }
                }
                helper.BatchExecuteBySql(lsb.ToArray(), lparam.ToArray());


                return index < 1;
            }
            catch (Exception e)
            {

                return false;
            }

        }
        private static MemoryStream SaveFile(List<A_Model> list, out string FileName)
        {
            FileName = "";
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            //  worksheet.Cells[2, 0].Value = "买方：TCL海外电子（惠州）有限公司";
            //   worksheet.Cells[4, 0].Value = "卖方：惠州市华星光电技术有限公司";
            Hashtable rl = new Hashtable();

            for (int i = 0; i < list.Count; i++)
            {
                if (i == 0)
                {
                    worksheet.Cells[0, 1].Value = "Material Code";
                    worksheet.Cells[0, 2].Value = "MAKTX_EN";
                    worksheet.Cells[0, 3].Value = "MAKTX_ZH";

                    worksheet.Cells[0, 4].Value = "ODF";
                    worksheet.Cells[0, 5].Value = "WorkCenter";
                    worksheet.Cells[0, 6].Value = "Demand";
                    worksheet.Cells[0, 7].Value = "Issued";
                    worksheet.Cells[0, 8].Value = "Issued Rate";

                    worksheet.Cells[i + 1, 1].Value = list[i].MATNR;
                    worksheet.Cells[i + 1, 2].Value = list[i].MAKTX_EN;
                    worksheet.Cells[i + 1, 3].Value = list[i].MAKTX_ZH;

                    worksheet.Cells[i + 1, 4].Value = list[i].ODF;
                    worksheet.Cells[i + 1, 5].Value = list[i].ARBPL;
                    worksheet.Cells[i + 1, 6].Value = list[i].Demand; ;
                    worksheet.Cells[i + 1, 7].Value = list[i].Issued; ;

                    double m = Convert.ToDouble(list[i].Issued) / Convert.ToDouble(list[i].Demand);
                    worksheet.Cells[i + 1, 8].Value = OptPerct((((m > 1 ? 1 : m) * 100).ToString() + "%"));
                }
                else
                {
                    worksheet.Cells[i + 1, 1].Value = list[i].MATNR;
                    worksheet.Cells[i + 1, 2].Value = list[i].MAKTX_EN;
                    worksheet.Cells[i + 1, 3].Value = list[i].MAKTX_ZH;

                    worksheet.Cells[i + 1, 4].Value = list[i].ODF;
                    worksheet.Cells[i + 1, 5].Value = list[i].ARBPL;
                    worksheet.Cells[i + 1, 6].Value = list[i].Demand; ;
                    worksheet.Cells[i + 1, 7].Value = list[i].Issued; ;
                    double m = Convert.ToDouble(list[i].Issued) / Convert.ToDouble(list[i].Demand);
                    worksheet.Cells[i + 1, 8].Value = OptPerct((((m > 1 ? 1 : m) * 100).ToString() + "%"));
                }


            }
            FileName = "Details.xls";
            workbook.FileFormat = FileFormatType.Excel97To2003;
            workbook.FileFormat = FileFormatType.Xlsx;

            return workbook.SaveToStream();

        }
        private static bool SynPICK_LIST(string sap_con, SqlHelper helper, DataTable dtW)
        {
            RfcDestination SapDest = null;
            SapDest = SapDesInit.GetDest(sap_con);

            RfcRepository RfcRep = SapDest.Repository;

            DataTable dt = new DataTable();

            DataTable dtFac = helper.GetTableBySql("SELECT Organization_Code FROM [dbo].[Base_Organization] WHERE ParentId IN(SELECT Organization_ID FROM [dbo].[Base_Organization] WHERE Organization_Code = '0988')");

            IRfcFunction GetWo = RfcRep.CreateFunction("ZCMP_UM_WM_PICK_LIST");//   
            IRfcTable IT_SELVAL = GetWo.GetTable("IT_SELVAL");

            foreach (DataRow drFac in dtFac.Rows)
            {
                IRfcStructure ZSPP_CMP_SELVAL1 = RfcRep.GetStructureMetadata("ZSPP_CMP_SELVAL").CreateStructure();
                ZSPP_CMP_SELVAL1.SetValue("FIELD", "WERKS");
                ZSPP_CMP_SELVAL1.SetValue("VAL_L", drFac[0].ToString());
                IT_SELVAL.Insert(ZSPP_CMP_SELVAL1);
            }
            foreach (DataRow item in dtW.Rows)
            {
                IRfcStructure ZSPP_CMP_SELVAL1 = RfcRep.GetStructureMetadata("ZSPP_CMP_SELVAL").CreateStructure();
                ZSPP_CMP_SELVAL1.SetValue("FIELD", "AUFNR");
                ZSPP_CMP_SELVAL1.SetValue("VAL_L", item["WorkOrderNo"].ToString());
                IT_SELVAL.Insert(ZSPP_CMP_SELVAL1);
            }

            GetWo.SetValue("IT_SELVAL", IT_SELVAL);

            GetWo.Invoke(SapDest);
            IRfcTable Wo = GetWo.GetTable("ET_TAB");
            dt.TableName = "TMSA_WM_PICK_LIST_ISSUERPT";
            dt.Columns.Add(new DataColumn("ID"));
            dt.Columns.Add(new DataColumn("PICKNO"));
            dt.Columns.Add(new DataColumn("ITEM"));
            dt.Columns.Add(new DataColumn("IDNRK"));
            dt.Columns.Add(new DataColumn("POSNR_B"));
            dt.Columns.Add(new DataColumn("LGORT"));
            dt.Columns.Add(new DataColumn("RSNUM"));
            dt.Columns.Add(new DataColumn("RSPOS"));
            dt.Columns.Add(new DataColumn("AUFNR"));
            dt.Columns.Add(new DataColumn("POSNR"));
            dt.Columns.Add(new DataColumn("GSTRP"));
            dt.Columns.Add(new DataColumn("KDAUF"));
            dt.Columns.Add(new DataColumn("KDPOS"));
            dt.Columns.Add(new DataColumn("BSTKD"));
            dt.Columns.Add(new DataColumn("CDATE"));
            dt.Columns.Add(new DataColumn("LGORT_K"));
            dt.Columns.Add(new DataColumn("BDTER"));
            dt.Columns.Add(new DataColumn("ARBPL"));
            dt.Columns.Add(new DataColumn("MAKTX1"));
            dt.Columns.Add(new DataColumn("LTXA1"));
            dt.Columns.Add(new DataColumn("MATNR"));
            dt.Columns.Add(new DataColumn("MAKTX"));
            dt.Columns.Add(new DataColumn("WERKS"));
            dt.Columns.Add(new DataColumn("DWERK"));
            dt.Columns.Add(new DataColumn("GAMNG"));
            dt.Columns.Add(new DataColumn("BDMNG"));
            dt.Columns.Add(new DataColumn("MEINS"));
            dt.Columns.Add(new DataColumn("SORTF"));
            dt.Columns.Add(new DataColumn("POTX1"));
            dt.Columns.Add(new DataColumn("ENMNG"));
            dt.Columns.Add(new DataColumn("LABST"));
            dt.Columns.Add(new DataColumn("BLQTY"));
            dt.Columns.Add(new DataColumn("DFLAG"));
            dt.Columns.Add(new DataColumn("PRT"));
            dt.Columns.Add(new DataColumn("IQTY"));
            dt.Columns.Add(new DataColumn("PQTY"));
            dt.Columns.Add(new DataColumn("EFLAG"));
            dt.Columns.Add(new DataColumn("DJMNG"));
            StringBuilder upsql = new StringBuilder();
            for (int i = 0; i < Wo.Count; i++)
            {
                Wo.CurrentIndex = i;
                DataRow dr = dt.NewRow();
                dr["ID"] = Guid.NewGuid().ToString();
                dr["PICKNO"] = Wo.GetString("PICKNO");
                dr["ITEM"] = Wo.GetString("ITEM");
                dr["IDNRK"] = Wo.GetString("IDNRK");
                dr["POSNR_B"] = Wo.GetString("POSNR_B");
                dr["LGORT"] = Wo.GetString("LGORT");
                dr["RSNUM"] = Wo.GetString("RSNUM");
                dr["RSPOS"] = Wo.GetString("RSPOS");
                dr["AUFNR"] = Wo.GetString("AUFNR");
                dr["POSNR"] = Wo.GetString("POSNR");
                dr["GSTRP"] = Wo.GetString("GSTRP");
                dr["KDAUF"] = Wo.GetString("KDAUF");
                dr["KDPOS"] = Wo.GetString("KDPOS");
                dr["BSTKD"] = Wo.GetString("BSTKD");
                dr["CDATE"] = Wo.GetString("CDATE");
                dr["LGORT_K"] = Wo.GetString("LGORT_K");
                dr["BDTER"] = Wo.GetString("BDTER");
                dr["ARBPL"] = !string.IsNullOrEmpty(Wo.GetString("ARBPL")) && Wo.GetString("ARBPL").StartsWith("TMSAFA") ? "TMSAFA" : Wo.GetString("ARBPL");
                dr["MAKTX1"] = Wo.GetString("MAKTX1");
                dr["LTXA1"] = Wo.GetString("LTXA1");
                dr["MATNR"] = Wo.GetString("MATNR");
                dr["MAKTX"] = Wo.GetString("MAKTX");
                dr["WERKS"] = Wo.GetString("WERKS");
                dr["DWERK"] = Wo.GetString("DWERK");
                dr["GAMNG"] = Wo.GetString("GAMNG");
                dr["BDMNG"] = Wo.GetString("BDMNG");
                dr["MEINS"] = Wo.GetString("MEINS");
                dr["SORTF"] = Wo.GetString("SORTF");
                dr["POTX1"] = Wo.GetString("POTX1");
                dr["ENMNG"] = Wo.GetString("ENMNG");
                dr["LABST"] = Wo.GetString("LABST");
                dr["BLQTY"] = Wo.GetString("BLQTY");
                dr["DFLAG"] = Wo.GetString("DFLAG");
                dr["PRT"] = Wo.GetString("PRT");
                dr["IQTY"] = Wo.GetString("IQTY");
                dr["PQTY"] = Wo.GetString("PQTY");
                dr["EFLAG"] = Wo.GetString("EFLAG");
                dr["DJMNG"] = Wo.GetString("DJMNG");
                dt.Rows.Add(dr);
            }
            helper.ExecuteBySql(new StringBuilder("DELETE TMSA_WM_PICK_LIST_ISSUERPT"), null);
            helper.MsSqlBulkCopyData(dt);
            //helper.ExecuteBySql(new StringBuilder("EXEC [dbo].[pro_UpZCMP_GI_LIST_RPT]"), null);
            return true;
        } 
        #endregion
    }
}

public class IssueModel
{
    public string ODF { get; set; }
    public string MQty { get; set; }
    public string CQty { get; set; }
    public string CRate { get; set; }
}
public class A_Model
{

    public string ODF { get; set; }
    public string MATNR { get; set; }
    public string ARBPL { get; set; }
    public string MAKTX_EN { get; set; }
    public string MAKTX_ZH { get; set; }
    public string Filt { get; set; }
    public string Issued { get; set; }
    public string Demand { get; set; }
}
public class CalModel
{
    public string ODF { get; set; }
    public string ARBPL { get; set; }
    public string RATES { get; set; }

}

public class FilterModel {
    public string ODF { get; set; }
    public string WORKCENTER { get; set; }
}