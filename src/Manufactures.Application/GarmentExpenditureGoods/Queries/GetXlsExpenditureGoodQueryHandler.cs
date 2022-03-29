using ExtCore.Data.Abstractions;
using Infrastructure.Domain.Queries;
using Infrastructure.External.DanLirisClient.Microservice;
using Infrastructure.External.DanLirisClient.Microservice.HttpClientService;
using Infrastructure.External.DanLirisClient.Microservice.MasterResult;
using Manufactures.Domain.GarmentFinishingOuts.Repositories;
using Manufactures.Domain.GarmentSewingOuts.Repositories;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using static Infrastructure.External.DanLirisClient.Microservice.MasterResult.CostCalculationGarmentDataProductionReport;
using static Infrastructure.External.DanLirisClient.Microservice.MasterResult.HOrderDataProductionReport;
using Manufactures.Domain.GarmentExpenditureGoods.Repositories;
using System.IO;
using System.Data;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Manufactures.Domain.GarmentPreparings.Repositories;
using Manufactures.Domain.GarmentCuttingIns.Repositories;
using System.Net.Http;
using System.Text;
using System.Globalization;

namespace Manufactures.Application.GarmentExpenditureGoods.Queries
{
	public class GetXlsExpenditureGoodQueryHandler : IQueryHandler<GetXlsExpenditureGoodQuery, MemoryStream>
	{
		protected readonly IHttpClientService _http;
		private readonly IStorage _storage;
		private readonly IGarmentExpenditureGoodRepository garmentExpenditureGoodRepository;
		private readonly IGarmentExpenditureGoodItemRepository garmentExpenditureGoodItemRepository;
		private readonly IGarmentPreparingRepository garmentPreparingRepository;
		private readonly IGarmentPreparingItemRepository garmentPreparingItemRepository;
		private readonly IGarmentCuttingInRepository garmentCuttingInRepository;
        private readonly IGarmentCuttingInItemRepository garmentCuttingInItemRepository;
        private readonly IGarmentCuttingInDetailRepository garmentCuttingInDetailRepository;

        public GetXlsExpenditureGoodQueryHandler(IStorage storage, IServiceProvider serviceProvider)
		{
			_storage = storage;
			garmentExpenditureGoodRepository = storage.GetRepository<IGarmentExpenditureGoodRepository>();
			garmentExpenditureGoodItemRepository = storage.GetRepository<IGarmentExpenditureGoodItemRepository>();
			garmentPreparingRepository = storage.GetRepository<IGarmentPreparingRepository>();
			garmentPreparingItemRepository = storage.GetRepository<IGarmentPreparingItemRepository>();
			garmentCuttingInRepository = storage.GetRepository<IGarmentCuttingInRepository>();
            garmentCuttingInItemRepository = storage.GetRepository<IGarmentCuttingInItemRepository>();
            garmentCuttingInDetailRepository = storage.GetRepository<IGarmentCuttingInDetailRepository>();
            _http = serviceProvider.GetService<IHttpClientService>();
		}

		class monitoringView
		{
			public string expenditureGoodNo { get; internal set; }
			public string expenditureGoodType { get; internal set; }
			public DateTimeOffset expenditureDate { get; internal set; }
			public string buyer { get; internal set; }
			public string roNo { get; internal set; }
			public string buyerArticle { get; internal set; }
			public string colour { get; internal set; }
			public string name { get; internal set; }
            public string unitname { get; internal set; }
			public double qty { get; internal set; }
			public string invoice { get; internal set; }
			public decimal price { get; internal set; }
			public double fc { get; set; }
		}

        public async Task<PEBResult> GetDataPEB(List<string> invoice, string token)
        {
            PEBResult pEB = new PEBResult();

            var listInvoice = string.Join(",", invoice.Distinct());
            var stringcontent = new StringContent(JsonConvert.SerializeObject(listInvoice), Encoding.UTF8, "application/json");

            var garmentProductionUri = CustomsDataSettings.Endpoint + $"customs-reports/getPEB";
            var httpResponse = await _http.SendAsync(HttpMethod.Get, garmentProductionUri, token, stringcontent);



            if (httpResponse.IsSuccessStatusCode)
            {
                var contentString = await httpResponse.Content.ReadAsStringAsync();
                Dictionary<string, object> content = JsonConvert.DeserializeObject<Dictionary<string, object>>(contentString);
                var dataString = content.GetValueOrDefault("data").ToString();

                var listdata = JsonConvert.DeserializeObject<List<PEBResultViewModel>>(dataString);

                foreach (var i in listdata)
                {
                    pEB.data.Add(i);
                }
                //garmentProduct.data = listdata;



            }

            return pEB;
        }

        public async Task<CostCalculationGarmentDataProductionReport> GetDataCostCal(List<string> ro, string token)
		{
			CostCalculationGarmentDataProductionReport costCalculationGarmentDataProductionReport = new CostCalculationGarmentDataProductionReport();

			var listRO = string.Join(",", ro.Distinct());
            var stringcontent = new StringContent(JsonConvert.SerializeObject(listRO), Encoding.UTF8, "application/json");
            var costCalculationUri = SalesDataSettings.Endpoint + $"cost-calculation-garments/data";
			//var httpResponse = await _http.GetAsync(costCalculationUri, token);

            var httpResponse = await _http.SendAsync(HttpMethod.Get, costCalculationUri, token, stringcontent);

            var freeRO = new List<string>();

			if (httpResponse.IsSuccessStatusCode)
			{
				var contentString = await httpResponse.Content.ReadAsStringAsync();
				Dictionary<string, object> content = JsonConvert.DeserializeObject<Dictionary<string, object>>(contentString);
				var dataString = content.GetValueOrDefault("data").ToString();
				var listData = JsonConvert.DeserializeObject<List<CostCalViewModel>>(dataString);

				foreach (var item in ro)
				{
					var data = listData.SingleOrDefault(s => s.ro == item);
					if (data != null)
					{
						costCalculationGarmentDataProductionReport.data.Add(data);
					}
					else
					{
						freeRO.Add(item);
					}
				}
			}

			 
			return costCalculationGarmentDataProductionReport;
		}
		 
		class ViewFC
		{
			public string RO { get; internal set; }
			public double FC { get; internal set; }
			public int Count { get; internal set; }
		}
		class ViewBasicPrices
		{
			public string RO { get; internal set; }
			public decimal BasicPrice { get; internal set; }
			public int Count { get; internal set; }
		}
		public async Task<MemoryStream> Handle(GetXlsExpenditureGoodQuery request, CancellationToken cancellationToken)
		{
			DateTimeOffset dateFrom = new DateTimeOffset(request.dateFrom);
			dateFrom.AddHours(7);
			DateTimeOffset dateTo = new DateTimeOffset(request.dateTo);
			dateTo = dateTo.AddHours(7);


			var QueryRo = (from a in garmentExpenditureGoodRepository.Query
						   where a.UnitId == (request.unit == 0 ? a.UnitId : request.unit) && a.ExpenditureDate >= dateFrom && a.ExpenditureDate <= dateTo
						   select a.RONo).Distinct().ToList();

			List<string> _ro = new List<string>();
			foreach (var item in QueryRo)
			{
				_ro.Add(item);
			}
			var _unitName = (from a in garmentPreparingRepository.Query
							 where a.UnitId == request.unit
							 select a.UnitName).FirstOrDefault();
			CostCalculationGarmentDataProductionReport costCalculation = await GetDataCostCal(_ro, request.token);
			GarmentMonitoringExpenditureGoodListViewModel listViewModel = new GarmentMonitoringExpenditureGoodListViewModel();
			List<GarmentMonitoringExpenditureGoodDto> monitoringDtos = new List<GarmentMonitoringExpenditureGoodDto>();
			var sumbasicPrice = (from a in garmentPreparingRepository.Query
								 join b in garmentPreparingItemRepository.Query on a.Identity equals b.GarmentPreparingId
								 where /*(request.ro == null || (request.ro != null && request.ro != "" && a.RONo == request.ro)) &&*/
								 a.UnitId == (request.unit == 0 ? a.UnitId : request.unit)
                                 select new { a.RONo, b.BasicPrice })
						.GroupBy(x => new { x.RONo }, (key, group) => new ViewBasicPrices
						{
							RO = key.RONo,
							BasicPrice = Convert.ToDecimal(group.Sum(s => s.BasicPrice)),
							Count = group.Count()
						});

			var sumFCs = (from a in garmentCuttingInRepository.Query
						  where /*(request.ro == null || (request.ro != null && request.ro != "" && a.RONo == request.ro)) && */ a.CuttingType == "Main Fabric" &&
						 a.UnitId == (request.unit == 0 ? a.UnitId : request.unit) && a.CuttingInDate <= dateTo
                          join b in garmentCuttingInItemRepository.Query on a.Identity equals b.CutInId
                          join c in garmentCuttingInDetailRepository.Query on b.Identity equals c.CutInItemId
                          select new { a.FC, a.RONo, FCs = Convert.ToDouble(c.CuttingInQuantity * a.FC), c.CuttingInQuantity })
                       .GroupBy(x => new { x.RONo }, (key, group) => new ViewFC
                       {
                           RO = key.RONo,
                           FC = group.Sum(s => (s.FCs)),
                           Count = group.Sum(s => s.CuttingInQuantity)
                       });
            var Query = from a in (from aa in garmentExpenditureGoodRepository.Query
								   where aa.UnitId == (request.unit == 0 ? aa.UnitId : request.unit) && aa.ExpenditureDate >= dateFrom && aa.ExpenditureDate <= dateTo
								   select aa)
						join b in garmentExpenditureGoodItemRepository.Query on a.Identity equals b.ExpenditureGoodId
						where a.UnitId == (request.unit == 0 ? a.UnitId : request.unit) && a.ExpenditureDate >= dateFrom && a.ExpenditureDate <= dateTo
						select new monitoringView { fc = (from aa in sumFCs where aa.RO == a.RONo select aa.FC / aa.Count).FirstOrDefault(),
                            price = Convert.ToDecimal((from aa in sumbasicPrice where aa.RO == a.RONo select aa.BasicPrice / aa.Count).FirstOrDefault()),
                            buyer = (from cost in costCalculation.data where cost.ro == a.RONo select cost.buyerCode).FirstOrDefault(), buyerArticle = a.BuyerCode + " " + a.Article,
                            roNo = a.RONo, expenditureDate = a.ExpenditureDate, expenditureGoodNo = a.ExpenditureGoodNo, expenditureGoodType = a.ExpenditureType, invoice = a.Invoice,
                            colour = b.Description, qty = b.Quantity, name = (from cost in costCalculation.data where cost.ro == a.RONo select cost.comodityName).FirstOrDefault(), unitname = a.UnitName };



			var querySum = Query.ToList().GroupBy(x => new { x.fc, x.buyer, x.buyerArticle, x.roNo, x.expenditureDate, x.expenditureGoodNo, x.expenditureGoodType, x.invoice, x.colour, x.name, x.unitname }, (key, group) => new
			{
				ros = key.roNo,
				buyer = key.buyerArticle,
				expenditureDates = key.expenditureDate,
				qty = group.Sum(s => s.qty),
				expendituregoodNo = key.expenditureGoodNo,
				expendituregoodTypes = key.expenditureGoodType,
				color = key.colour,
				price = group.Sum(s => s.price),
				buyerC = key.buyer,
				names = key.name,
                unitname = key.unitname,
				invoices = key.invoice,
				fcs = key.fc

			}).OrderBy(s => s.expendituregoodNo);

            var Pebs = await GetDataPEB(querySum.Select(x => x.invoices).ToList(), request.token);

            foreach (var item in querySum)
			{
                var peb = Pebs.data.FirstOrDefault(x => x.BonNo.Trim() == item.invoices);
                GarmentMonitoringExpenditureGoodDto dto = new GarmentMonitoringExpenditureGoodDto
				{
					roNo = item.ros,
					buyerArticle = item.buyer,
					expenditureGoodType = item.expendituregoodTypes,
					expenditureGoodNo = item.expendituregoodNo,
					expenditureDate = item.expenditureDates,
                    pebDate = peb == null ? new DateTime(1970,1,1) : peb.BCDate,
                    qty = item.qty,
					colour = item.color,
					name = item.names,
                    unitname = item.unitname,
					invoice = item.invoices,
					price = Math.Round(Convert.ToDecimal(Convert.ToDouble(Math.Round(item.price, 2)) * Math.Round(item.fcs, 2)), 2),
					buyerCode = item.buyerC,
					nominal = Math.Round( Convert.ToDecimal(item.qty) * Convert.ToDecimal(Convert.ToDouble(Math.Round(item.price, 2)) * Math.Round(item.fcs, 2)),2)

				};
				monitoringDtos.Add(dto);
			}
			var data = from a in monitoringDtos
					   where a.qty > 0
					   select a;
			monitoringDtos = data.ToList();
			double qty = 0;
			decimal nominal = 0;
			foreach (var item in data)
			{
				qty += item.qty;
				nominal += item.nominal;

			}
			GarmentMonitoringExpenditureGoodDto dtos = new GarmentMonitoringExpenditureGoodDto
			{
				roNo ="",
				buyerArticle = "",
				expenditureGoodType = "",
				expenditureGoodNo = "",
				expenditureDate=null, 
                pebDate = null,
				qty = qty,
				colour = "",
				name = "",
                unitname = "",
				invoice = "",
				price = 0,
				buyerCode = "",
				nominal = nominal

			};
			monitoringDtos.Add(dtos);
			listViewModel.garmentMonitorings = monitoringDtos;
			var reportDataTable = new DataTable();
			reportDataTable.Columns.Add(new DataColumn() { ColumnName = "NO BON", DataType = typeof(string) });
			reportDataTable.Columns.Add(new DataColumn() { ColumnName = "TIPE PENGELUARAN", DataType = typeof(string) });
			reportDataTable.Columns.Add(new DataColumn() { ColumnName = "TGL", DataType = typeof(string) });
			reportDataTable.Columns.Add(new DataColumn() { ColumnName = "TGL PEB", DataType = typeof(string) });
			reportDataTable.Columns.Add(new DataColumn() { ColumnName = "RO", DataType = typeof(string) });
			reportDataTable.Columns.Add(new DataColumn() { ColumnName = "BUYER & ARTICLE", DataType = typeof(string) });
			reportDataTable.Columns.Add(new DataColumn() { ColumnName = "COLOUR", DataType = typeof(string) });
			reportDataTable.Columns.Add(new DataColumn() { ColumnName = "NAMA", DataType = typeof(string) });
			reportDataTable.Columns.Add(new DataColumn() { ColumnName = "UNIT", DataType = typeof(string) });
			reportDataTable.Columns.Add(new DataColumn() { ColumnName = "HARGA (PCS)", DataType = typeof(decimal) });
			reportDataTable.Columns.Add(new DataColumn() { ColumnName = "QTY", DataType = typeof(double) });
			reportDataTable.Columns.Add(new DataColumn() { ColumnName = "NOMINAL", DataType = typeof(double) });
			reportDataTable.Columns.Add(new DataColumn() { ColumnName = "INVOICE", DataType = typeof(string) });
			int counter = 5;
			if (listViewModel.garmentMonitorings.Count > 0)
			{
				foreach (var report in listViewModel.garmentMonitorings)
				{
                    
                    string pebDate = report.pebDate.GetValueOrDefault() == new DateTime(1970, 1, 1) || report.pebDate.GetValueOrDefault().ToString("dd MMM yyyy") == "01 Jan 0001" ? "-" : report.pebDate.GetValueOrDefault().ToString("dd MMM yyy");
                    //Console.WriteLine(pebDate);
                    reportDataTable.Rows.Add(report.expenditureGoodNo, report.expenditureGoodType, report.expenditureDate.GetValueOrDefault().ToOffset(new TimeSpan(7,0,0)).ToString("dd MMM yyy"), pebDate,
                    report.roNo, report.buyerArticle, report.colour, report.name, report.unitname, report.price, report.qty,report.nominal, report.invoice);
					counter++;
                    //Console.WriteLine(counter);
                }
			}
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet 1");
				worksheet.Cells["A" + 5 + ":L" + 5 + ""].Style.Font.Bold = true;
				worksheet.Cells["A1"].Value = "Report Barang Jadi "; worksheet.Cells["A" + 1 + ":L" + 1 + ""].Merge = true;
				worksheet.Cells["A2"].Value = "Periode " + dateFrom.ToString("dd-MM-yyyy") + " s/d " + dateTo.ToString("dd-MM-yyyy");
				worksheet.Cells["A3"].Value = "Konfeksi " + _unitName;
				worksheet.Cells["A" + 1 + ":L" + 1 + ""].Merge = true;
				worksheet.Cells["A" + 2 + ":L" + 2 + ""].Merge = true;
				worksheet.Cells["A" + 3 + ":L" + 3 + ""].Merge = true;
				worksheet.Cells["A" + 1 + ":L" + 3 + ""].Style.Font.Size = 15;
				worksheet.Cells["A" + 1 + ":L" + 5 + ""].Style.Font.Bold = true;
				worksheet.Cells["A" + 1 + ":L" + 5 + ""].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				worksheet.Cells["A" + 1 + ":L" + 5 + ""].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
				worksheet.Cells["A5"].LoadFromDataTable(reportDataTable, true);
				worksheet.Column(8).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
				worksheet.Cells["I" + 2 + ":K" + counter + ""].Style.Numberformat.Format = "#,##0.00";
				worksheet.Column(9).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
				worksheet.Cells["J" + 2 + ":J" + counter + ""].Style.Numberformat.Format = "#,##0.00";
				worksheet.Cells["A" + 5 + ":L" + counter + ""].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
				worksheet.Cells["A" + 5 + ":L" + counter + ""].Style.Border.Top.Style = ExcelBorderStyle.Thin;
				worksheet.Cells["A" + 5 + ":L" + counter + ""].Style.Border.Left.Style = ExcelBorderStyle.Thin;
				worksheet.Cells["A" + 5 + ":L" + counter + ""].Style.Border.Right.Style = ExcelBorderStyle.Thin;
				worksheet.Cells["I" + (counter) + ":J" + (counter) + ""].Style.Font.Bold = true;
				worksheet.Cells["A" + 5 + ":L" + 5 + ""].Style.Font.Bold = true;
				var stream = new MemoryStream();
				if (request.type != "bookkeeping")
				{
					worksheet.Cells["A" + (counter) + ":I" + (counter) + ""].Merge = true;

					worksheet.Column(9).Hidden = true;
				}else
				{
					worksheet.Cells["A" + (counter) + ":H" + (counter) + ""].Merge = true;
				}
				package.SaveAs(stream);

				return stream;
			}
		}
	}
}
