using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using WebAPI.Models;
using static WebAPI.Constants.Constant;

namespace WebAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FileController : ControllerBase
    {
        [HttpPost]
        public IActionResult ImportFile(IFormFile file)
        {
            using (var stream = new MemoryStream())
            {
                file.CopyToAsync(stream);
                using (ExcelPackage package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["CV-HTV-Nguyen Anh Tu"];
                    int totalRows = worksheet.Dimension.Rows;
                    int totalColums = worksheet.Dimension.Columns;
                    List<string> list = new List<string>();
                    for (int row = 1; row <= 60; row++)
                    {
                        for (int col = 1; col <= totalRows; col++)
                        {
                            if (!string.IsNullOrEmpty(worksheet.Cells[row, col].Value?.ToString().Trim()))
                            {
                                list.Add(worksheet.Cells[row, col].Value?.ToString().Trim());
                            }
                        }
                    }
                    var person = GetPerson(list);
                    var listCategory = GetListCategory(list);
                    var listTechnology = GetListTechnology(list);
                    var listWorkHistory = GetListWorkHistory(list);
                    var listEducation = GetListEducation(list);
                    return Ok(listEducation);
                }
            }
        }
        #region Xu Ly Person
        /// <summary>
        /// Get Person
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        private Person GetPerson(List<string> list)
        {
            Person person = new Person
            {
                FullName = list[1],
                Gender = FormatGender(list[3]),
                YearOfBirth = FormatYearOfBirth(list[5].Trim()),
                Description = list[7]
            };
            return person;
        }

        /// <summary>
        /// Format Gender
        /// </summary>
        /// <param name="genderStr"></param>
        /// <returns></returns>
        private eGender FormatGender(string genderStr)
        {
            int gerder = 2;
            if (!String.IsNullOrEmpty(genderStr))
            {
                switch (genderStr)
                {
                    case "Male":
                        gerder = 0;
                        break;
                    case "Female":
                        gerder = 1;
                        break;
                    case "Sexless":
                        gerder = 2;
                        break;
                    default:
                        break;
                }
            }
            return (Constants.Constant.eGender)gerder;
        }

        /// <summary>
        /// Format Year Of Birth
        /// </summary>
        /// <param name="yearOfBirthStr"></param>
        /// <returns></returns>
        private DateTime FormatYearOfBirth(string yearOfBirthStr)
        {
            DateTime yearOfBirth = new DateTime();
            if (!String.IsNullOrEmpty(yearOfBirthStr))
            {
                if (yearOfBirthStr.Contains("-"))
                {
                    int x = 0;
                    for (int i = 0; i < yearOfBirthStr.Length; i++)
                    {
                        if (yearOfBirthStr[i].Equals("-"))
                        {
                            x++;
                        }
                    }
                    if (x == 1)
                    {
                        String[] arrListStr = yearOfBirthStr.Split('-');
                        if (!String.IsNullOrEmpty(arrListStr[0]) && !String.IsNullOrEmpty(arrListStr[1]))
                        {
                            int month = Convert.ToInt32(arrListStr[0]);
                            int year = Convert.ToInt32(arrListStr[1]);
                            yearOfBirth = new DateTime(year, month, 01);
                        }
                    }
                    else if (x == 2)
                    {
                        String[] arrListStr = yearOfBirthStr.Split('-');
                        if (!String.IsNullOrEmpty(arrListStr[0]) && !String.IsNullOrEmpty(arrListStr[1]) && !String.IsNullOrEmpty(arrListStr[2]))
                        {
                            int day = Convert.ToInt32(arrListStr[0]);
                            int month = Convert.ToInt32(arrListStr[1]);
                            int year = Convert.ToInt32(arrListStr[2]);
                            yearOfBirth = new DateTime(year, month, day);
                        }
                    }
                }
                else
                {
                    int year = Convert.ToInt32(yearOfBirthStr);
                    yearOfBirth = new DateTime(year, 01, 01);
                }
            }
            return yearOfBirth;
        }
        #endregion
        #region Xu Ly Skill
        /// <summary>
        /// Get List Category
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        private List<Category> GetListCategory(List<string> list)
        {
            List<Category> listCategory = new List<Category>();
            int indexSkill = list.IndexOf("SKILLS");
            int indexWorkHistory = list.IndexOf("WORKING HISTORY");
            for (int i = indexSkill; i < indexWorkHistory; i++)
            {
                if ((i - (indexSkill + 1)) % 2 == 0)
                {
                    Category category = new Category
                    {
                        Name = list[i]
                    };
                    listCategory.Add(category);
                }
            }

            return listCategory;
        }

        /// <summary>
        /// Get List Technology
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        private List<Technology> GetListTechnology(List<string> list)
        {
            List<Technology> listTechnology = new List<Technology>();
            int indexSkill = list.IndexOf("SKILLS");
            int indexWorkHistory = list.IndexOf("WORKING HISTORY");
            for (int i = indexSkill; i < indexWorkHistory; i++)
            {
                if (((i - (indexSkill)) % 2 == 0) && i > indexSkill)
                {
                    var listItemTechnology = SplitTechnology(list[i].Trim());
                    if (listItemTechnology != null)
                    {
                        foreach (var item in listItemTechnology)
                        {
                            Technology technology = new Technology
                            {
                                Name = item
                            };
                            listTechnology.Add(technology);
                        }
                    }

                }
            }

            return listTechnology;
        }
        /// <summary>
        /// Split Technology
        /// </summary>
        /// <param name="technoloryStr"></param>
        /// <returns></returns>
        private List<string> SplitTechnology(string technoloryStr)
        {
            List<string> listTechnology = new List<string>();
            if (!String.IsNullOrEmpty(technoloryStr))
            {
                if (technoloryStr.Contains(","))
                {
                    String[] arrListStr = technoloryStr.Split(',');
                    for (int i = 0; i < arrListStr.Length; i++)
                    {
                        listTechnology.Add(arrListStr[i]);
                    }
                }
            }
            return listTechnology;
        }
        #endregion
        #region Xu Ly WorkHistory
        private List<WorkHistoryInfo> GetListWorkHistory(List<string> list)
        {
            List<WorkHistoryInfo> listWorkHistory = new List<WorkHistoryInfo>();
            int indexWorkHistory = list.IndexOf("WORKING HISTORY");
            int indexEducation = list.IndexOf("EDUCATION");
            for (int i = indexWorkHistory + 4; i < indexEducation-4; i = i + 4)
            {
                WorkHistoryInfo workHistoryInfo = new WorkHistoryInfo
                {
                    OrderIndex = Convert.ToInt32(list[i+1]),
                    StartDate = FormatStartDateWorkHistory(list[i+2]),
                    EndDate = FormatEndDateWorkHistory(list[i+2]),
                    CompanyName = list[i + 3],
                    Position = list[i + 4]
                };
                listWorkHistory.Add(workHistoryInfo);
            }
            return listWorkHistory;
            
        }
         
        private int GetMonth(string str)
        {
            int result = 1;
            switch (str)
            {
                case "Jan":
                    result = 1;
                    break;
                case "Feb":
                    result = 2;
                    break;
                case "Mar":
                    result = 3;
                    break;
                case "Apr":
                    result = 4;
                    break;
                case "May":
                    result = 5;
                    break;
                case "Jun":
                    result = 6;
                    break;
                case "Jul":
                    result = 7;
                    break;
                case "Aug":
                    result = 8;
                    break;
                case "Sep":
                    result = 9;
                    break;
                case "Oct":
                    result = 10;
                    break;
                case "Nov":
                    result = 11;
                    break;
                case "Dec":
                    result = 12;
                    break;
                default:
                    break;
            }
            return result;
        }

        private DateTime FormatEndDateWorkHistory(string dateStr)
        {
            DateTime endDate = new DateTime();
            if (!String.IsNullOrEmpty(dateStr))
            {
                string[] arrListStr = dateStr.Split("-");
                if (!String.IsNullOrEmpty(arrListStr[0]))
                {
                    if (arrListStr[0].Contains("Now"))
                    {
                        endDate = DateTime.Now;
                    }
                    else
                    {
                        string[] arrListStartDate = arrListStr[0].Split(" ");
                        if (!String.IsNullOrEmpty(arrListStartDate[0]) && !String.IsNullOrEmpty(arrListStartDate[1]))
                        {
                            int month = GetMonth(arrListStartDate[0]);
                            int year = Convert.ToInt32(arrListStartDate[1]);
                            endDate = new DateTime(year, month, 01);
                        }
                    }
                    
                }
            }
            return endDate;
        }
        private DateTime FormatStartDateWorkHistory(string dateStr)
        {
            DateTime strartDate = new DateTime();
            if (!String.IsNullOrEmpty(dateStr))
            {
                string[] arrListStr = dateStr.Split("-");
                if (!String.IsNullOrEmpty(arrListStr[1]))
                {
                    string[] arrListStartDate = arrListStr[1].TrimStart().Split(" ");
                    if(!String.IsNullOrEmpty(arrListStartDate[0]) && !String.IsNullOrEmpty(arrListStartDate[1]))
                    {
                        int month = GetMonth(arrListStartDate[0]);
                        int year = Convert.ToInt32(arrListStartDate[1]);
                        strartDate = new DateTime(year, month, 01);
                    }
                }
            }
            return strartDate;
        }
        #endregion
        #region Xu Ly Education
        private List<EducationInfo> GetListEducation(List<string> list)
        {
            List<EducationInfo> listEducation = new List<EducationInfo>();
            int indexEducation = list.IndexOf("EDUCATION");
            int indexCertificate = list.IndexOf("CERTIFICATION");
            for(int i = indexEducation; i < indexCertificate - 2; i=i+2)
            {
                EducationInfo educationInfo = new EducationInfo
                {
                    StartDate = FormatStartDateEducation(GetDateEducation(list[i + 1])),
                    EndDate = FormatGetEndDateEducation(GetDateEducation(list[i + 1])),
                    CollegeName = GetCollegeName(list[i+1]),
                    Major = GetMajorEducation(list[i + 2])
                };
                listEducation.Add(educationInfo);
            }
            return listEducation;
        }

        private DateTime FormatStartDateEducation(string str)
        {
            DateTime startDate = new DateTime();
            if (!String.IsNullOrEmpty(str))
            {
                string[] arrStr = str.Split("-");
                if (arrStr != null && arrStr.Length > 1)
                {
                    if (!String.IsNullOrEmpty(arrStr[1]))
                    {
                        int year = Convert.ToInt32(arrStr[1]);
                        startDate = new DateTime(year, 01, 01);
                    } 
                }
            }
            return startDate;
        }

        private DateTime FormatGetEndDateEducation(string str)
        {
            DateTime endDate = new DateTime();
            if (!String.IsNullOrEmpty(str))
            {
                string[] arrStr = str.Split("-");
                if (arrStr != null && arrStr.Length > 1)
                {
                    if (!String.IsNullOrEmpty(arrStr[0]))
                    {
                        int year = Convert.ToInt32(arrStr[0]);
                        endDate = new DateTime(year, 01, 01);
                    }
                }
            }
            return endDate;
        }

        private string GetCollegeName(string str)
        {
            string result = null;
            if (!String.IsNullOrEmpty(str))
            {
                if (str.Contains("|"))
                {
                    string[] arrStr = str.Split("|");
                    if(arrStr != null && arrStr.Length >1)
                    {
                        result = arrStr[1].Trim();
                    }
                }
            }
            return result;
        }

        private string GetDateEducation(string str)
        {
            string result = null;
            if (!String.IsNullOrEmpty(str))
            {
                if (str.Contains("|"))
                {
                    string[] arrStr = str.Split("|");
                    if (arrStr != null && arrStr.Length > 1)
                    {
                        result = arrStr[0].Trim();
                    }
                }
            }
            return result;
        }

        private string GetMajorEducation(string str)
        {
            string result = null;
            if (!String.IsNullOrEmpty(str))
            {
                if (str.Contains(":"))
                {
                    string[] arrStr = str.Split(":");
                    if (arrStr != null && arrStr.Length > 1)
                    {
                        result = arrStr[1].Trim();
                    }
                }
            }
            return result;
        }
        #endregion
        #region Xu ly Certification
        private List<CertificateInfo> GetListCertifiate(List<string> list)
        {
            List<CertificateInfo> listCertificate = new List<CertificateInfo>();
            int indexCertificate = list.IndexOf("CERTIFICATION");
            int indexProject = list.IndexOf("PROJECTS");
            for(int i = indexCertificate; i < indexProject; i++)
            {
                CertificateInfo certificateInfo = new CertificateInfo
                {
                    //StartDate
                };
                list.Add(certificateInfo);
            }
            return listCertificate;
        }
        private string GetNameCertifiate(string str)
        {
            string result = null;
            if (!String.IsNullOrEmpty(str))
            {
                
            }
            return result;
        }
        private DateTime GetStartDateCertifiate(string str)
        {
            List<DateTime> listDate = new List<DateTime>();
            if (!String.IsNullOrEmpty(str))
            {
                string[] arrStr = str.Split("|");
                if (arrStr.Any())
                {
                    if (arrStr[0].Contains(-))
                    {

                    }
                    else
                    {
                        if (!String.IsNullOrEmpty(arrStr[0]))
                        {
                            int year = Convert.ToInt32(arrStr[0]);
                            DateTime startDate = new DateTime(year, 01, 01);
                            listDate.Add(startDate);
                        }
                    }
                    
                }
            }
            return listDate;
        }
        #endregion
    }
}
