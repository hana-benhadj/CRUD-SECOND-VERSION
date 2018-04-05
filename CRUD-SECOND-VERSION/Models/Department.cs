using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CRUD_SECOND_VERSION.Models
{
    public class Department
    {
        [Key]
        public int IdDep { get; set; }
        [Required]
        public string Name { get; set; }
        [Display(Name = "Employee")]
        public virtual ICollection<Employee> Emp { get; set; }
        [NotMapped]
        public HttpPostedFileBase Image { get; set; }
        [NotMapped]
        public string ImageUrl { get {
               return IdDep.ToString()+".jpg";
            } }
    }
}