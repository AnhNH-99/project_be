using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using WebAPI.Models;
using static WebAPI.Constants.Constant;

namespace WebAPI.HandlingFiles
{
    public static class ImportFileHandling
    {
        public static async Task<List<String>> GetInformationCV(IFormFile file)
        {
            List<string> list = new List<string>();
            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
                using (ExcelPackage package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["CV-HTV-Nguyen Anh Tu"];
                    int totalRows = worksheet.Dimension.Rows;
                    int totalColums = worksheet.Dimension.Columns;
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
                    int indexSkill = list.IndexOf("SKILLS");
                    int indexWorkHistory = list.IndexOf("WORKING HISTORY");
                    int indexEducation = list.IndexOf("EDUCATION");
                    int indexCertificate = list.IndexOf("CERTIFICATION");
                    int indexProject = list.IndexOf("PROJECTS");
                    int indexEnd = list.Count();
                }
            }
            return list;
        }

        #region Handling Person
        /// <summary>
        /// Get Person
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static Person GetPerson(List<string> list)
        {
            Hashtable hashtable = FormatYear(list[5]);
            Person person = new Person
            {
                StaffId = "",
                FullName = list[1],
                Email = "",
                Location = eLocation.HAN,
                Avatar = "",
                Description = list[7],
                Phone = "",
                YearOfBirth = (DateTime?)hashtable["StartDate"],
                Gender = FormatGender(list[3]),
                CreatedBy = WebAPI.Helpers.HttpContext.CurrentUser,
                UpdatedBy = WebAPI.Helpers.HttpContext.CurrentUser,
                Status = true

            };
            return person;
        }

        /// <summary>
        /// Format Gender
        /// </summary>
        /// <param name="genderStr"></param>
        /// <returns></returns>
        public static eGender FormatGender(string genderStr)
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
        #region Handling Skill
        /// <summary>
        /// Get List Category
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static List<Category> GetListCategory(List<string> list)
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
                        Name = list[i],
                        Technologies = new List<Technology>(),
                    };
                    List<string> listTechnology = SplitTechnology(GetInformation(list[i + 1]));
                    foreach (var item in listTechnology)
                    {
                        Technology technology = new Technology
                        {
                            Name = item
                        };
                        category.Technologies.Add(technology);
                    }
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
        public static List<Technology> GetListTechnology(List<string> list)
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
        public static List<string> SplitTechnology(string technoloryStr)
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
        #region Handling WorkHistory
        /// <summary>
        /// Get List WorkHistory
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static List<WorkHistoryInfo> GetListWorkHistory(List<string> list)
        {
            List<WorkHistoryInfo> listWorkHistory = new List<WorkHistoryInfo>();
            int indexWorkHistory = list.IndexOf("WORKING HISTORY");
            int indexEducation = list.IndexOf("EDUCATION");
            for (int i = indexWorkHistory + 4; i < indexEducation - 4; i = i + 4)
            {
                Hashtable hashtable = FormatYearAndMonth(list[i + 2]);
                WorkHistoryInfo workHistoryInfo = new WorkHistoryInfo
                {
                    OrderIndex = Convert.ToInt32(list[i + 1]),
                    StartDate = (DateTime?)hashtable["StartDate"],
                    EndDate = (DateTime?)hashtable["EndDate"],
                    CompanyName = list[i + 3],
                    Position = list[i + 4]
                };
                listWorkHistory.Add(workHistoryInfo);
            }
            return listWorkHistory;

        }
        #endregion
        #region Handling Education
        /// <summary>
        /// Get List Education
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static List<EducationInfo> GetListEducation(List<string> list)
        {
            List<EducationInfo> listEducation = new List<EducationInfo>();
            int indexEducation = list.IndexOf("EDUCATION");
            int indexCertificate = list.IndexOf("CERTIFICATION");
            int index = 0;
            for (int i = indexEducation; i < indexCertificate - 2; i = i + 2)
            {
                index++;
                EducationInfo educationInfo = new EducationInfo();
                string dateEducation = GetDate(list[i + 1]);
                Hashtable hashtable = FormatYear(dateEducation);
                if (hashtable.Count > 1)
                {
                    educationInfo = new EducationInfo
                    {
                        StartDate = (DateTime)hashtable["StartDate"],
                        EndDate = (DateTime)hashtable["EndDate"],
                        CollegeName = GetCollegeName(list[i + 1]),
                        Major = GetInformation(list[i + 2]),
                        OrderIndex = index
                    };
                }
                else
                {
                    educationInfo = new EducationInfo
                    {
                        StartDate = (DateTime)hashtable["StartDate"],
                        EndDate = null,
                        CollegeName = GetCollegeName(list[i + 1]),
                        Major = GetInformation(list[i + 2]),
                        OrderIndex = index
                    };
                }
                listEducation.Add(educationInfo);
            }
            return listEducation;
        }

        /// <summary>
        /// Get ColleName
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string GetCollegeName(string str)
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
        #endregion
        #region Handling Certification
        /// <summary>
        /// Get List Certifiate
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static List<CertificateInfo> GetListCertifiate(List<string> list)
        {
            List<CertificateInfo> listCertificate = new List<CertificateInfo>();
            int indexCertificate = list.IndexOf("CERTIFICATION");
            int indexProject = list.IndexOf("PROJECTS");
            int index = 0;
            for (int i = indexCertificate + 1; i < indexProject; i++)
            {
                index++;
                CertificateInfo certificateInfo = new CertificateInfo();
                string certificateStr = list[i];
                string dateCertificate = GetDate(certificateStr);
                Hashtable hashtable = FormatYear(dateCertificate);
                if (hashtable.Count > 1)
                {
                    certificateInfo = new CertificateInfo
                    {
                        StartDate = (DateTime)hashtable["StartDate"],
                        EndDate = (DateTime?)hashtable["EndDate"],
                        Name = GetNameAndProviderCertifiate(certificateStr)["Name"].ToString(),
                        Provider = GetNameAndProviderCertifiate(certificateStr)["Provider"].ToString(),
                        OrderIndex = index
                    };
                }
                else
                {
                    certificateInfo = new CertificateInfo
                    {
                        StartDate = (DateTime)hashtable["StartDate"],
                        EndDate = null,
                        Name = GetNameAndProviderCertifiate(certificateStr)["Name"].ToString(),
                        Provider = GetNameAndProviderCertifiate(certificateStr)["Provider"].ToString(),
                        OrderIndex = index
                    };
                }
                listCertificate.Add(certificateInfo);
            }
            return listCertificate;
        }
        /// <summary>
        /// Get Name And Provider Certifiate
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static Hashtable GetNameAndProviderCertifiate(string str)
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
                        if (arrStrNameAndProvider.Any() && arrStrNameAndProvider.Length > 1)
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
        #region Handling Project
        /// <summary>
        /// Get List Project
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static List<Project> GetListProject(List<string> list)
        {
            int indexProject = list.IndexOf("PROJECTS");
            int indexEnd = list.Count();
            List<Project> listProject = new List<Project>();
            for (int i = indexProject + 4; i < indexEnd - 8; i = i + 8)
            {
                string str = list[i + 2];
                Hashtable hashDate = FormatYearAndMonth(list[i + 2]);
                Project project = new Project
                {
                    OrderIndex = (String.IsNullOrEmpty(list[i + 1])) ? 0 : Convert.ToInt32(list[i + 1]),
                    StartDate = (DateTime)hashDate["StartDate"],
                    EndDate = (DateTime)hashDate["EndDate"],
                    Position = list[i + 3],
                    Name = list[i + 4],
                    Description = GetInformation(list[i + 5]),
                    Responsibilities = GetInformation(list[i + 6]),
                    TeamSize = (String.IsNullOrEmpty(GetInformation((list[i + 7])))) ? 1 : Convert.ToInt32(GetInformation(list[i + 7])),
                    Technologies = new List<Technology>()
                };
                List<string> listTechnology = SplitTechnology(GetInformation(list[i + 8]));
                foreach (var item in listTechnology)
                {
                    Technology technology = new Technology
                    {
                        Name = item
                    };
                    project.Technologies.Add(technology);
                }
                listProject.Add(project);
            }
            return listProject;
        }
        #endregion
        #region Handling Date
        /// <summary>
        /// Get Date
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string GetDate(string str)
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
        /// <summary>
        /// Get Month
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static int GetMonth(string str)
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

        /// <summary>
        /// FormatYearAndMonth
        /// </summary>
        /// <param name="dateStr"></param>
        /// <returns></returns>
        public static Hashtable FormatYearAndMonth(string dateStr)
        {
            Hashtable hashtable = new Hashtable();
            if (!String.IsNullOrEmpty(dateStr))
            {
                string[] arrListStr = dateStr.Trim().Split("-");
                if (!String.IsNullOrEmpty(arrListStr[0]) && !String.IsNullOrEmpty(arrListStr[1]))
                {
                    if (arrListStr[0].Contains("Now"))
                    {
                        hashtable.Add("EndDate", DateTime.Now);
                    }
                    else
                    {
                        string[] arrListEndtDate = arrListStr[0].Trim().Split(" ");
                        if (!String.IsNullOrEmpty(arrListEndtDate[0]) && !String.IsNullOrEmpty(arrListEndtDate[1]))
                        {
                            int month = GetMonth(arrListEndtDate[0]);
                            int year = Convert.ToInt32(arrListEndtDate[1]);
                            DateTime endDate = new DateTime(year, month, 01);
                            hashtable.Add("EndDate", endDate);
                        }
                    }
                    string[] arrListStarttDate = arrListStr[1].Trim().Split(" ");
                    if (!String.IsNullOrEmpty(arrListStarttDate[0]) && !String.IsNullOrEmpty(arrListStarttDate[1]))
                    {
                        int month = GetMonth(arrListStarttDate[0]);
                        int year = Convert.ToInt32(arrListStarttDate[1]);
                        DateTime startDate = new DateTime(year, month, 01);
                        hashtable.Add("StartDate", startDate);
                    }
                }
            }
            return hashtable;
        }

        /// <summary>
        /// Format Year
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static Hashtable FormatYear(string str)
        {
            Hashtable hashtable = new Hashtable();
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
                            hashtable.Add("EndDate", endDate);
                            int yearStartDate = Convert.ToInt32(arrStr[1]);
                            DateTime startDate = new DateTime(yearStartDate, 01, 01);
                            hashtable.Add("StartDate", startDate);
                        }
                    }
                }
                else
                {
                    int year = Convert.ToInt32(str);
                    DateTime dateTime = new DateTime(year, 01, 01);
                    hashtable.Add("StartDate", dateTime);
                }
            }
            return hashtable;
        }
        #endregion
        #region Handling Information
        /// <summary>
        /// Get Infomation
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string GetInformation(string str)
        {
            string result = null;
            if (!String.IsNullOrEmpty(str))
            {
                if (str.Contains(":"))
                {
                    string[] arrStr = str.Trim().Split(":");
                    if (arrStr.Length > 1)
                    {
                        result = (String.IsNullOrEmpty(arrStr[1])) ? null : arrStr[1].Trim();
                    }
                }
            }
            return result;
        }
        #endregion

    }
}
