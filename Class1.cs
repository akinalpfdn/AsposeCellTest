using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellTest
{
	class Class1
	{
		public void AsposeDeneme()
		{
			Workbook wb = new Workbook(@"C:\Users\akina\Downloads\asposeDeneme.xlsx");
			int sheetNumber = 0;
			Worksheet ws = wb.Worksheets[sheetNumber];
			//ws.Cells.Merge(6, 1, 2, 2);
			AutoFitterOptions options = new AutoFitterOptions();

			options.AutoFitMergedCells = true;

			int pageCount = (int)Math.Ceiling(123412 / 84.0);
			ws.AutoFitRows(options);
			ws.Cells.SetRowHeight(0, 12.75);
			ws.Cells[0, 16].PutValue("DENEME1");//Dokuman No
			ws.Cells[1, 16].PutValue(DateTime.Now);//Yayin Tarihi
			ws.Cells[4, 16].PutValue(pageCount);//Sayfa No
			ws.ViewType = ViewType.PageBreakPreview;
			using (Stream respStream = new MemoryStream())
			{

				wb.Save(respStream, Aspose.Cells.SaveFormat.Xlsx);
				respStream.Seek(0, SeekOrigin.Begin);
				//WriteToResponse(respStream, "denemeAsposeOut" + ".xlsx");
			}
		}
	}
}
