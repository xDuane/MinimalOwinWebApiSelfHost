using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;
using System.Net.Http;
using MinimalOwinWebApiSelfHost.Models;

using Microsoft.Office.Interop.Excel;

// Add these usings:
using System.Data.Entity;
using System.Reflection;

using System.Runtime.Caching;

namespace MinimalOwinWebApiSelfHost.Controllers
{
    //[Authorize(Roles = "Admin")]
    public class CompaniesController : ApiController
    {
        ApplicationDbContext dbContext = new ApplicationDbContext();
        private DateTime _cacheTime = DateTime.Today;
        private int _cacheDays = 1;

        public IEnumerable<Company> GetList(string zip)
        {
            var cache = MemoryCache.Default;
            var _companies = cache["_companies"] as IEnumerable<Company>;
            var _cacheTime = cache["_cacheTime"] as DateTime?;
            var _cacheDays = cache["_cacheDays"] as int?;

            if (_companies == null || _cacheTime == null || _cacheDays == null || _cacheTime < DateTime.Now.AddDays(-_cacheDays.Value))
            {
                _companies = RDS.Classes.FileProcess.loadFile();
            }

            return _companies.Where(z => z.zip == zip).ToList();
        }

        public string[] GetRange(string range, Worksheet excelWorksheet)
        {
            var workingRangeCells = excelWorksheet.get_Range(range, Type.Missing);

            var array = (System.Array)workingRangeCells.Cells.Value2;
            string[] arrayS = array.OfType<object>().Select(o => o.ToString()).ToArray(); ; // this.ConvertToStringArray(array);

            return arrayS;
        }

        public async Task<Company> Get(int id = -1)
        {
            var company = await dbContext.Companies.FirstOrDefaultAsync(c => c.Id == id);
            if (company == null)
            {
                throw new HttpResponseException(
                    System.Net.HttpStatusCode.NotFound);
            }
            return company;
        }

        public async Task<IHttpActionResult> Post(Company company)
        {
            if (company == null)
            {
                return BadRequest("Argument Null");
            }
            var companyExists = await dbContext.Companies.AnyAsync(c => c.Id == company.Id);

            if (companyExists)
            {
                return BadRequest("Exists");
            }

            dbContext.Companies.Add(company);
            await dbContext.SaveChangesAsync();
            return Ok();
        }

        public async Task<IHttpActionResult> Put(Company company)
        {
            if (company == null)
            {
                return BadRequest("Argument Null");
            }
            var existing = await dbContext.Companies.FirstOrDefaultAsync(c => c.Id == company.Id);

            if (existing == null)
            {
                return NotFound();
            }

            existing.Name = company.Name;
            await dbContext.SaveChangesAsync();
            return Ok();
        }

        public async Task<IHttpActionResult> Delete(int id)
        {
            var company = await dbContext.Companies.FirstOrDefaultAsync(c => c.Id == id);
            if (company == null)
            {
                return NotFound();
            }
            dbContext.Companies.Remove(company);
            await dbContext.SaveChangesAsync();
            return Ok();
        }
    }
}
