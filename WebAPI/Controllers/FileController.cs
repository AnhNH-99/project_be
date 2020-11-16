using System;
using System.Collections;
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
                    var listCertificate = GetListCertifiate(list);
                    return Ok(listCertificate);
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
                YearOfBirth = FormatYear(list[5]).FirstOrDefault(),
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
            for (int i = indexWorkHistory + 4; i < indexEducation - 4; i = i + 4)
            {
                WorkHistoryInfo workHistoryInfo = new WorkHistoryInfo
                {
                    OrderIndex = Convert.ToInt32(list[i + 1]),
                    StartDate = FormatYearAndMonth(list[i + 2])[1],
                    EndDate = FormatYearAndMonth(list[i + 2])[0],
                    CompanyName = list[i + 3],
                    Position = list[i + 4]
                };
                listWorkHistory.Add(workHistoryInfo);
            }
            return listWorkHistory;

        }
        #endregion
        #region Xu Ly Education
        private List<EducationInfo> GetListEducation(List<string> list)
        {
            List<EducationInfo> listEducation = new List<EducationInfo>();
            int indexEducation = list.IndexOf("EDUCATION");
            int indexCertificate = list.IndexOf("CERTIFICATION");
            for (int i = indexEducation; i < indexCertificate - 2; i = i + 2)
            {
                EducationInfo educationInfo = new EducationInfo();
                string dateEducation = GetDate(list[i + 1]);
                if (FormatYear(dateEducation).Count > 1)
                {
                    educationInfo = new EducationInfo
                    {
                        StartDate = FormatYear(dateEducation)[1],
                        EndDate = FormatYear(dateEducation)[0],
                        CollegeName = GetCollegeName(list[i + 1]),
                        Major = GetMajorEducation(list[i + 2])
                    };
                }
                else
                {
                    educationInfo = new EducationInfo
                    {
                        StartDate = FormatYear(dateEducation)[0],
                        EndDate = null,
                        CollegeName = GetCollegeName(list[i + 1]),
                        Major = GetMajorEducation(list[i + 2])
                    };
                }
                listEducation.Add(educationInfo);
            }
            return listEducation;
        }

        private string GetCollegeName(string str)
        {
            string result = null;
            if (!String.IsNullOrEmpty(str))
            {
                if (str.Contains("|"))
                {
                    string[] arrStr = str.Split("|");
                    if (arrStr != null && arrStr.Length > 1)
                    {
                        result = arrStr[1].Trim();
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
            for (int i = indexCertificate+1; i < indexProject; i++)
            {
                CertificateInfo certificateInfo = new CertificateInfo();
                string certificateStr = list[i];
                string dateCertificate = GetDate(certificateStr);
                if (FormatYear(dateCertificate).Count > 1)
                {
                    certificateInfo = new CertificateInfo
                    {
                        StartDate = FormatYear(dateCertificate)[1],
                        EndDate = FormatYear(dateCertificate)[0],
                        Name = GetNameAndProviderCertifiate(certificateStr)["Name"].ToString(),
                        Provider = GetNameAndProviderCertifiate(certificateStr)["Provider"].ToString()
                    };
                }
                else
                {
                    certificateInfo = new CertificateInfo
                    {
                        StartDate = FormatYear(dateCertificate).FirstOrDefault(),
                        EndDate = null,
                        Name = GetNameAndProviderCertifiate(certificateStr)["Name"].ToString(),
                        Provider = GetNameAndProviderCertifiate(certificateStr)["Provider"].ToString()
                    };
                }
              listCertificate.Add(certificateInfo);
            }
            return listCertificate;
        }
        private Hashtable GetNameAndProviderCertifiate(string str)
        {

            Hashtable hashtable = new Hashtable();
            if (!String.IsNullOrEmpty(str))
            {
                if (str.Contains("-"))
                {
                    string[] arrStr = str.Trim().Split("|");
                    if (arrStr.Any() && arrStr.Length > 1)
                    {
                        string[] arrStrNameAndProvider = arrStr[1].Trim().Split("-");
                        if(arrStrNameAndProvider.Any()&& arrStrNameAndProvider.Length > 1)
                        {
                            hashtable.Add("Name", arrStrNameAndProvider[0]);
                            hashtable.Add("Provider", arrStrNameAndProvider[1]);
                        }
                    }
                }
            }
            return hashtable;
        }

        #endregion

        #region Format Date
        private string GetDate(string str)
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

        private List<DateTime> FormatYearAndMonth(string dateStr)
        {
            List<DateTime> listDateTime = new List<DateTime>();
            if (!String.IsNullOrEmpty(dateStr))
            {
                string[] arrListStr = dateStr.Trim().Split("-");
                if (!String.IsNullOrEmpty(arrListStr[0]) && !String.IsNullOrEmpty(arrListStr[1]))
                {
                    if (arrListStr[0].Contains("Now"))
                    {
                        listDateTime.Add(DateTime.Now);
                    }
                    else
                    {
                        string[] arrListEndtDate = arrListStr[0].Trim().Split(" ");
                        if (!String.IsNullOrEmpty(arrListEndtDate[0]) && !String.IsNullOrEmpty(arrListEndtDate[1]))
                        {
                            int month = GetMonth(arrListEndtDate[0]);
                            int year = Convert.ToInt32(arrListEndtDate[1]);
                            DateTime endDate = new DateTime(year, month, 01);
                            listDateTime.Add(endDate);
                        }
                    }
                    string[] arrListStarttDate = arrListStr[1].Trim().Split(" ");
                    if (!String.IsNullOrEmpty(arrListStarttDate[0]) && !String.IsNullOrEmpty(arrListStarttDate[1]))
                    {
                        int month = GetMonth(arrListStarttDate[0]);
                        int year = Convert.ToInt32(arrListStarttDate[1]);
                        DateTime startDate = new DateTime(year, month, 01);
                        listDateTime.Add(startDate);
                    }
                }
            }
            return listDateTime;
        }

        public List<DateTime> FormatYear(string str)
        {
            List<DateTime> listDateTime = new List<DateTime>();
            if (!String.IsNullOrEmpty(str))
            {
                if (str.Contains("-"))
                {
                    string[] arrStr = str.Trim().Split("-");
                    if (arrStr.Any())
                    {
                        if (!String.IsNullOrEmpty(arrStr[0]) && !String.IsNullOrEmpty(arrStr[1]))
                        {
                            int yearEndDate = Convert.ToInt32(arrStr[0]);
                            DateTime endDate = new DateTime(yearEndDate, 01, 01);
                            listDateTime.Add(endDate);
                            int yearStartDate = Convert.ToInt32(arrStr[1]);
                            DateTime startDate = new DateTime(yearStartDate, 01, 01);
                            listDateTime.Add(startDate);
                        }
                    }
                }
                else
                {
                    int year = Convert.ToInt32(str);
                    DateTime dateTime = new DateTime(year, 01, 01);
                    listDateTime.Add(dateTime);
                }
            }
            return listDateTime;
        }
        #endregion
    }
}
