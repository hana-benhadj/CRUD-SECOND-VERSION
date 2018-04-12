using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using CRUD_SECOND_VERSION.Models;
using Excel;
using Microsoft.Reporting.WebForms;
using OfficeOpenXml;
using PagedList;
using PagedList.Mvc;

using  Excel = Microsoft.Office.Interop.Excel;

namespace CRUD_SECOND_VERSION.Controllers
{
    public class EmployeesController : Controller
    {
        private ApplicationDbContext db = new ApplicationDbContext();

        // GET: Employees
        public ActionResult Index(String search, string sortOrder, string currentFilter, int? page, int? pagesize)
        {
            var employees = db.Employees.Include(e => e.Department);
            //search
            //The search string value is received from the text box
            if (!String.IsNullOrEmpty(search))
            {
                employees = employees.Where(s => s.Name.Contains(search) ||
                s.Position.Contains(search) || s.Department.Name.Contains(search));
            }
            if (search != null)
            {
                page = 1;
                 
            }
            else
            {
                search = currentFilter;
            }
            List<SelectListItem> items = new List<SelectListItem>{          
            new SelectListItem{ Text="5", Value="5" },
             new SelectListItem{ Text="10", Value="10" },
            new SelectListItem{ Text="20", Value="20" },
             new SelectListItem{ Text="50", Value="50" },
            new SelectListItem{ Text="100", Value="100" }
        };

            ViewBag.CurrentFilter = search;
            ViewBag.CurrentSort = sortOrder;
            //The two ViewBag variables with the appropriate query string values
            //if the sortOrder parameter is null or empty, ViewBag.NameSortParm should be set to 
            //"name_desc"...  it should be set to an empty string.
            ViewBag.NameSortParm = String.IsNullOrEmpty(sortOrder) ? "name_desc" : "";
            ViewBag.PositionSortParm = String.IsNullOrEmpty(sortOrder) ? "position_desc" : "";
            var emp = from s in db.Employees
                      select s;
            //After you click the Name heading, employees are 
            //displayed in descending name order.
            switch (sortOrder)
            {
                case "name_desc":
                    employees = employees.OrderByDescending(s => s.Name);
                    break;
                case "position_desc":
                    employees = employees.OrderByDescending(s => s.Position);
                    break;

                default:
                    employees = employees.OrderBy(s => s.Name);
                    break;
            }
            //pagination
            // 1) installing the PagedList.Mvc NuGet package
            // PagedList.Mvc is one of many good paging and sorting packages for ASP.NET MVC
            //The NuGet PagedList.Mvc package automatically installs the PagedList 
            //package as a dependency.

            //2) This code adds a page parameter, a current sort order parameter, and a current
            //filter parameter to the method signature
            int pageSize = pagesize ?? 5;       
            int pageNumber = (page ?? 1);
            ViewBag.pagesize = new SelectList(items, "value", "Text", pagesize);
            ViewBag.CurrentPageSize = pageSize;
  
            return View(employees.ToPagedList(pageNumber, pageSize));
        }

            // GET: Employees/Details/5
            public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Employee employee = db.Employees.Find(id);
            if (employee == null)
            {
                return HttpNotFound();
            }
            return View(employee);
        }

        // GET: Employees/Create
        public ActionResult Create()
        {
            ViewBag.IdDep = new SelectList(db.Departments, "IdDep", "Name");
            return View();
        }

        // POST: Employees/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,Name,Position,IdDep")] Employee employee)
        {
            //var query = db.Employees.Select(c => new SelectListItem
            //{
            //    Value = c.Department.ImageUrl.ToString(),
            //    Text = c.Department.Name               
            //});

            if (ModelState.IsValid)
            {
                db.Employees.Add(employee);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.IdDep = new SelectList(db.Departments,"IdDep", "Name", employee.IdDep);
            return View(employee);
           
        }

        // GET: Employees/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Employee employee = db.Employees.Find(id);
            if (employee == null)
            {
                return HttpNotFound();
            }
            ViewBag.IdDep = new SelectList(db.Departments, "IdDep", "Name", employee.IdDep);
            return View(employee);
        }

        // POST: Employees/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,Name,Position,IdDep")] Employee employee)
        {
            if (ModelState.IsValid)
            {
                db.Entry(employee).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.IdDep = new SelectList(db.Departments, "IdDep", "Name", employee.IdDep);
            return View(employee);
        }

        // GET: Employees/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Employee employee = db.Employees.Find(id);
            if (employee == null)
            {
                return HttpNotFound();
            }
            return View(employee);
        }

        // POST: Employees/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Employee employee = db.Employees.Find(id);
            db.Employees.Remove(employee);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        public ActionResult View ( int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Employee employee = db.Employees.Find(id);
            if (employee == null)
            {
                return HttpNotFound();
            }
            return PartialView(employee);
        }

        public ActionResult CreateEmpView(Employee employee)
        {

            if (ModelState.IsValid)
            {
                db.Employees.Add(employee);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.IdDep = new SelectList(db.Departments, "IdDep", "Name", employee.IdDep);
            return PartialView(employee);
           
        }
        public ActionResult DeleteEmpView(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Employee employee = db.Employees.Find(id);
            if (employee == null)
            {
                return HttpNotFound();
            }
            return PartialView(employee);
        }
      
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        // function ExportToExcel
        public void ExportToExcel()
        {
            var employees = db.Employees.Include(e => e.Department).ToList();
            ExcelPackage pck = new ExcelPackage();
            ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Report");
            ws.Cells["A1"].Value = "Communication";
            ws.Cells["B1"].Value = "Com1";

            ws.Cells["A2"].Value = "Report";
            ws.Cells["B2"].Value = "report1";

            ws.Cells["A3"].Value = "Date";
            ws.Cells["B2"].Value = string.Format("{0:dd MMMM yyyy} at {0:H: mm tt}", DateTimeOffset.Now);

            // start table 
            ws.Cells["A6"].Value = "ID";
            ws.Cells["B6"].Value = "Name";
           
            ws.Cells["D6"].Value = "Position";
            ws.Cells["E6"].Value = "Department Name";

            int rowStart = 7;
            foreach (var item in employees)
            {
                //if (item.Experience < 5)
                //{
                //    ws.Row(rowStart).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //    ws.Row(rowStart).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(string.Format("pink")));

                //}

                ws.Cells[string.Format("A{0}", rowStart)].Value = item.Id;
                ws.Cells[string.Format("B{0}", rowStart)].Value = item.Name;
                ws.Cells[string.Format("C{0}", rowStart)].Value = item.Position;
                ws.Cells[string.Format("D{0}", rowStart)].Value = item.Department.Name;
                rowStart++;
            }
            // end table
            ws.Cells["A:AZ"].AutoFitColumns();
            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment: filename=" + "ExcelReport.xlsx");
            Response.BinaryWrite(pck.GetAsByteArray());
            Response.End();


        }// end  ExportToExcel


        // GET: Departments/Details/5
        public ActionResult Upload()
        {
            return View();
        }


        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Upload(HttpPostedFileBase upload)
        {
            if (ModelState.IsValid)
            {

                if (upload != null && upload.ContentLength > 0)
                {
                    // ExcelDataReader works with the binary Excel file, so it needs a FileStream
                    // to get started. This is how we avoid dependencies on ACE or Interop:
                    Stream stream = upload.InputStream;

                    // We return the interface, so that
                    IExcelDataReader reader = null;


                    if (upload.FileName.EndsWith(".xls"))
                    {
                        reader = ExcelReaderFactory.CreateBinaryReader(stream);
                    }
                    else if (upload.FileName.EndsWith(".xlsx"))
                    {
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    }
                    else
                    {
                        ModelState.AddModelError("File", "This file format is not supported");
                        return View();
                    }

                    reader.IsFirstRowAsColumnNames = true;
                    DataSet result = reader.AsDataSet();
                    reader.Close();

                    return View(result.Tables[0]);
                }
                else
                {
                    ModelState.AddModelError("File", "Please Upload Your file");
                }
            }
            return View();
        }
        public ActionResult Reports(string ReportType)
        {
            LocalReport localReport = new LocalReport();
            localReport.ReportPath = Server.MapPath("~/Reports/Report.rdlc");
            ReportDataSource reportDataSource = new ReportDataSource();
            reportDataSource.Name = "EmployeeDataSet";
            reportDataSource.Value = db.Employees.ToList();
            localReport.DataSources.Add(reportDataSource);
            string mimeType;
            string encoding;
            string fileNameExtension;
            if (ReportType == "Excel")
            {
                fileNameExtension = "xlsx" ;
            }
            else if (ReportType == "Word")
            {
                fileNameExtension = "docx";
            }
            else if (ReportType == "PDF")
            {
                fileNameExtension = "pdf";
            }
            else if (ReportType == "Image")
            {
                fileNameExtension = "jpg";
            }
            string[] streams;
            Warning[] warnings;
            byte[] renderedByte;
            renderedByte = localReport.Render(ReportType, "", out mimeType, out encoding, out fileNameExtension, out streams, out warnings);
            Response.AddHeader("content-disposition", "attachment: filename=" + "Report." + fileNameExtension);
            return File(renderedByte, fileNameExtension);
        }
    }
}
