using System;
using System.Text.Json.Serialization;
using WebAPI.Models.Token;

namespace WebAPI.Models
{
    public class AccountInfo
    {
        /// <summary>
        /// USER_NO
        /// </summary>
        public int USER_NO { get; set; }

        /// <summary>
        /// USER_CD
        /// </summary>
        public string USER_CD { get; set; }

        /// <summary>
        /// USER_NAME
        /// </summary>
        public string USER_NAME { get; set; }

        /// <summary>
        /// PASSWORD
        /// </summary>
        [JsonIgnore]
        public string PASSWORD { get; set; }

        /// <summary>
        /// FULL_NAME
        /// </summary>
        public string FULL_NAME { get; set; }

        /// <summary>
        /// EMAIL
        /// </summary>
        public string EMAIL { get; set; }

        /// <summary>
        /// PHONE
        /// </summary>
        public string PHONE { get; set; }

        /// <summary>
        /// ADDRESS
        /// </summary>
        public string ADDRESS { get; set; }

        /// <summary>
        /// ROLE
        /// </summary>
        public string ROLE { get; set; }

        /// <summary>
        /// STATUS FLG
        /// </summary>
        public string STATUS_FLG { get; set; }

        /// <summary>
        /// CREATED AT
        /// </summary>
        public DateTime CREATED_AT { get; set; }

        /// <summary>
        /// CREATED BY
        /// </summary>
        public string CREATED_BY { get; set; }

        /// <summary>
        /// UPDATED AT
        /// </summary>
        public DateTime UPDATED_AT { get; set; }

        /// <summary>
        /// UPDATED BY
        /// </summary>
        public string UPDATED_BY { get; set; }

        /// <summary>
        /// RefreshToken
        /// </summary>
        [JsonIgnore]
        public RefreshToken RefreshToken { get; set; }

    }
}
