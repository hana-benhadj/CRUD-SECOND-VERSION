using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using CRUD_SECOND_VERSION.Models;
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
        public ActionResult Import()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Import(HttpPostedFileBase excelfile)
        {
            if (excelfile==null || excelfile.ContentLength == 0)
            {
                ViewBag.Error = "Please select an excel file <br>";
                return View();
            }
            else
            {
                if(excelfile.FileName.EndsWith("xls") || excelfile.FileName.EndsWith("xlsx"))
                {
                    string path = Server.MapPath("~/Content/" + excelfile.FileName);
                    if (System.IO.File.Exists(path))
                        System.IO.File.Delete(path);
                    excelfile.SaveAs(path);
                    //read data from excel file
                    Excel.Application application = new Excel.Application();
                   Excel.Workbook workbook = application.Workbooks.Open(path).ActiveSheet;
                   Excel.Worksheet worksheet = workbook.ActiveSheet;
                   Excel.Range range = worksheet.UsedRange;
                    List<Employee> employees = new List<Employee>();
                    for(int row=3; row < range.Rows.Count; row++)
                    {
                        Employee e = new Employee();
                        e.Name = (range.Cells[ row,1]).Text;
                        e.Position = (range.Cells[row, 2]).Text;
                        e.Department.Name = (range.Cells[row, 3]).Text;
                        employees.Add(e);
                    }
                    ViewBag.employees = employees;

                    return View("Success");
                }
                else
                {
                    ViewBag.Error = "File type is incorrect";
                    return View();
                }
            }

        }
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
