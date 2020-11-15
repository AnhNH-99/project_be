namespace WebAPI.Constants
{
    public class Constant
    {
        public static string CONNECT_ERROR = "Please specify the correct connection string.";
        public enum eLocation : byte
        {
            HAN,
            DN,
            HCM
        }
        public enum eGender : byte
        {
            Male,
            Female,
            Sexless
        }
    }
    
}
