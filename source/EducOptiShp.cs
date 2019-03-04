using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Threading;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace EducOpti.source
{
    class EducOptiShp
    {

       

        public  DataTable CalNewRelation(DataTable districtDS, DataTable  schoolDS, DataTable  relationDS) 
        {   
        //    DataTable relationDT = relationDS.Tables[0];
        //     DataTable  districtDT=  districtDS.Tables[0];
        //      DataTable  schoolDT= schoolDS.Tables[0];


            DataTable relationDT = relationDS;
             DataTable  districtDT=  districtDS;
             DataTable schoolDT = schoolDS;

            int districtNum=districtDT.Rows.Count;
            int schoolNum=schoolDT.Rows.Count;
            int relationNum=relationDT.Rows.Count;
          //  Console.WriteLine("districtNum:{0},schoolNum:{1},relationNum:{2}", districtNum, schoolNum, relationNum);
          //  Console.WriteLine("districtfirst:{0},schoolfirst:{1},relationfirst:{2}", districtDT.Rows[0][districtDT.Columns[0]].ToString()
          //  , schoolDT.Rows[0][schoolDT.Columns[0]].ToString(), relationDT.Rows[0][relationDT.Columns[0]].ToString());

            int[] isDis = new int[districtNum];//判断小区是否已经分配
            int[] isSchMax = new int[schoolNum];//学校是否已到最大
            int[] SchNum = new int[schoolNum];//学校包含人数
            int[] isSelect = new int[relationNum];//判断学校小区关系是否已经选择
           
            //对学校小区关系按关联得分排序
          

            DataTable sortDT = relationDT.Copy();

            DataView dv = relationDT.DefaultView;
            dv.Sort = string.Format("{0} Desc", relationDT.Columns[9]);//关联得分字段排序
           // Console.WriteLine("关联得分字段:{0}", relationDT.Rows[0][relationDT.Columns[9]].ToString());
            sortDT = dv.ToTable();

            //Console.WriteLine(" sortDT.Rows.Count:{0}", sortDT.Rows.Count);

            //Console.WriteLine("关联得分字段:{0}", sortDT.Rows[0][sortDT.Columns[9]].ToString());
            sortDT.Columns.Add("selectTime",typeof(int));

            for (int i = 0; i < relationNum;i++ )
            {
                int tempDisId = int.Parse(sortDT.Rows[i][sortDT.Columns[8]].ToString()) ;
                int tempSchId = int.Parse(sortDT.Rows[i][sortDT.Columns[7]].ToString());
                //Console.WriteLine("tempDisId:{0} tempSchId:{1}", tempDisId, tempSchId);
                //foreach (DataRow dr in districtDT.Rows)
                int findSet_XQ =0;
                int findSet_XX=0;
                for (int j = 0; j < districtNum;j++ )
                {
                    if (int.Parse(districtDT.Rows[j][districtDT.Columns[0]].ToString()) == tempDisId)
                    { findSet_XQ = j; break; }

                }
                for (int j = 0; j < schoolNum; j++)
                {
                    if (int.Parse(schoolDT.Rows[j][schoolDT.Columns[0]].ToString()) == tempSchId)
                    { findSet_XX = j; break; }
                }


                if (isDis[findSet_XQ]==1 || isSchMax[findSet_XX]==1)
                { continue; }
                else
                {
                    int tempXQRS =int.Parse(districtDT.Rows[findSet_XQ][districtDT.Columns[1]].ToString());
                    int tempMaxXXRS=int.Parse(schoolDT.Rows[findSet_XX][schoolDT.Columns[1]].ToString());
                    //Console.WriteLine("tempXQRS:{0} tempMaxXXRS:{1}", tempXQRS, tempMaxXXRS);
                    isDis[findSet_XQ] = 1;
                    SchNum[findSet_XX] += tempXQRS;
                     if  (SchNum[findSet_XX] >= tempMaxXXRS)
                     {  isSchMax[findSet_XX] = 1;}
                  
                    isSelect[i] = 1;

                   
                    sortDT.Rows[i][sortDT.Columns[12]] = 1;
                  
                }
            }
            List<int> disList=new List<int>();
            //还未选的小区按距离最近原则直接选择
           // Console.WriteLine("还未选的小区:");
            for (int i = 0; i < districtNum; i++)
            {
                if (isDis[i] == 0) { disList.Add(i);
               
                //Console.WriteLine("{0}", i);
                }
            }
            int count=disList.Count;
           
            for (int i = 0; i < count; i++)
            {     
            int a=disList[i];
            List<int> relationList = new List<int>();
            
          
           // Console.WriteLine("还未选的小区{0}有如下几条关系：", a);
                for (int j = 0; j < relationNum; j++)
                {
                    if (int.Parse(sortDT.Rows[j][sortDT.Columns[8]].ToString()) == a)
                    { relationList.Add(j);  }
                }

                int c = 0;
                double Mindist = double.MaxValue;
                int b = relationList.Count;
                for (int j = 0; j < b; j++)
                {
                    if (Mindist > double.Parse(sortDT.Rows[relationList[j]][sortDT.Columns[6]].ToString()))
                    {
                        Mindist = double.Parse(sortDT.Rows[relationList[j]][sortDT.Columns[6]].ToString());
                        c = relationList[j];
                        //Console.WriteLine("还未选的小区{0}最短：{1}：", a,c);
                    }
                }

                isSelect[c] = 1;

                
                sortDT.Rows[c][sortDT.Columns[12]] = 2;
            }
           DataTable dt1=new DataTable();
           dt1 = sortDT.Copy();//必须需要表结构
           dt1.Rows.Clear();
            for (int j = 0; j < relationNum; j++)
            {

                if (isSelect[j]==1) {
                    
                   
                    dt1.ImportRow(sortDT.Rows[j]);
                   
                   //Console.WriteLine(" dt1.Rows.Count:{0},j:{1}", dt1.Rows.Count, j);
                   //Console.WriteLine("dt1第一字段:{0}", dt1.Rows[dt1.Rows.Count - 1][dt1.Columns[0]].ToString());
                }
                
            }
            //Console.WriteLine(" dt1.Rows.CountTotal:{0}", dt1.Rows.Count);
            //DataSet ds1=new DataSet();
           // ds1.Tables.Add(dt1.Copy());
            return dt1;

        }

      /// <summary>  
        /// 将excel导入到datatable  
        /// </summary>  
        /// <param name="filePath">excel路径</param>  
        /// <param name="isColumnName">第一行是否是列名</param>  
        /// <returns>返回datatable</returns>  
        public  DataTable ExcelToDataTable(string filePath, bool isColumnName)  
        {  
            DataTable dataTable = null;  
            FileStream fs = null;  
            DataColumn column = null;  
            DataRow dataRow = null;  
            IWorkbook workbook = null;  
            ISheet sheet = null;  
            IRow row = null;  
            ICell cell = null;  
            int startRow = 0;  
            try  
            {  
                using (fs = File.OpenRead(filePath))  
                {  
                    // 2007版本  
                    if (filePath.IndexOf(".xlsx") > 0)  
                        workbook = new XSSFWorkbook(fs);  
                    // 2003版本  
                    else if (filePath.IndexOf(".xls") > 0)  
                        workbook = new HSSFWorkbook(fs);  
  
                    if (workbook != null)  
                    {  
                        sheet = workbook.GetSheetAt(0);//读取第一个sheet，当然也可以循环读取每个sheet  
                        dataTable = new DataTable();  
                        if (sheet != null)  
                        {  
                            int rowCount = sheet.LastRowNum;//总行数  
                            if (rowCount > 0)  
                            {  
                                IRow firstRow = sheet.GetRow(0);//第一行  
                                int cellCount = firstRow.LastCellNum;//列数  
  
                                //构建datatable的列  
                                if (isColumnName)  
                                {  
                                    startRow = 1;//如果第一行是列名，则从第二行开始读取  
                                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)  
                                    {  
                                        cell = firstRow.GetCell(i);  
                                        if (cell != null)  
                                        {  
                                            if (cell.StringCellValue != null)  
                                            {  
                                                column = new DataColumn(cell.StringCellValue);  
                                                dataTable.Columns.Add(column);  
                                            }  
                                        }  
                                    }  
                                }  
                                else  
                                {  
                                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)  
                                    {  
                                        column = new DataColumn("column" + (i + 1));  
                                        dataTable.Columns.Add(column);  
                                    }  
                                }  
  
                                //填充行  
                                for (int i = startRow; i <= rowCount; ++i)  
                                {  
                                    row = sheet.GetRow(i);  
                                    if (row == null) continue;  
  
                                    dataRow = dataTable.NewRow();  
                                    for (int j = row.FirstCellNum; j < cellCount; ++j)  
                                    {  
                                        cell = row.GetCell(j);                                          
                                        if (cell == null)  
                                        {  
                                            dataRow[j] = "";  
                                        }  
                                        else  
                                        {  
                                            //CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)  
                                            switch (cell.CellType)  
                                            {  
                                                case CellType.Blank:  
                                                    dataRow[j] = "";  
                                                    break;  
                                                case CellType.Numeric:  
                                                    short format = cell.CellStyle.DataFormat;  
                                                    //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理  
                                                    if (format == 14 || format == 31 || format == 57 || format == 58)  
                                                        dataRow[j] = cell.DateCellValue;  
                                                    else  
                                                        dataRow[j] = cell.NumericCellValue;  
                                                    break;  
                                                case CellType.String:  
                                                    dataRow[j] = cell.StringCellValue;  
                                                    break;  
                                            }  
                                        }  
                                    }  
                                    dataTable.Rows.Add(dataRow);  
                                }  
                            }  
                        }  
                    }  
                }  
                return dataTable;  
            }  
            catch (Exception)  
            {  
                if (fs != null)  
                {  
                    fs.Close();  
                }  
                return null;  
            }  
        }  
    
        public  bool DataTableToExcel(DataTable dt,string str1)
        {
            bool result = false;
            IWorkbook workbook = null;
            FileStream fs = null;
            IRow row = null;
            ISheet sheet = null;
            ICell cell = null;
            try
            {
                if (dt != null && dt.Rows.Count > 0)
                {
                    workbook = new HSSFWorkbook();
                    sheet = workbook.CreateSheet("Sheet0");//创建一个名称为Sheet0的表  
                    int rowCount = dt.Rows.Count;//行数  
                    int columnCount = dt.Columns.Count;//列数  

                    //设置列头  
                    row = sheet.CreateRow(0);//excel第一行设为列头  
                    for (int c = 0; c < columnCount; c++)
                    {
                        cell = row.CreateCell(c);
                        cell.SetCellValue(dt.Columns[c].ColumnName);
                    }

                    //设置每行每列的单元格,  
                    for (int i = 0; i < rowCount; i++)
                    {
                        row = sheet.CreateRow(i + 1);
                        for (int j = 0; j < columnCount; j++)
                        {
                            cell = row.CreateCell(j);//excel第二行开始写入数据  
                            cell.SetCellValue(dt.Rows[i][j].ToString());
                        }
                    }
                    using (fs = File.OpenWrite(str1))
                    {
                        workbook.Write(fs);//向打开的这个xls文件中写入数据  
                        result = true;
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                return false;
            }
        }  
    }

    
}
