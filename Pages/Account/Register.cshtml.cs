using Mayntech___Individual_Solution.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Identity.Client.Platforms.Features.DesktopOs.Kerberos;
using System.ComponentModel.DataAnnotations;
using System.Data.SqlClient;
using System.Drawing;

namespace Mayntech___Individual_Solution.Pages.Account
{
    public class RegisterModel : PageModel
    {
        [BindProperty]
        public Registration registration { get; set; }

        public String ErrorMsg = "";

        private readonly IConfiguration _config;


        public RegisterModel( IConfiguration config)
        {
            
            _config = config;   
        }

        public void OnGet()
        {
            
        }

        //public async Task<IActionResult> OnPostAsync()
        //{
        //    if (!ModelState.IsValid) return Page();

        //    if (true)
        //    {

        //    }
        //}

        public RedirectToPageResult OnPost()
        {
            var environment = _config.GetValue<string>("ASPNETCORE_ENVIRONMENT");

            string connectionString = _config["connectionString"];

            string DuplicateUserName = null;
            string DuplicateUserEmail = null;

            

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                try
                {
                    String sql = "SELECT UserName FROM UserList WHERE UserName ='" + registration.UserName + "'";
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            reader.Read();
                            DuplicateUserName = reader.GetString(0);
                        }
                    }


                }
                catch (Exception)
                {

                    
                }
                try
                {
                    String sqlEmail = "SELECT Email FROM UserList WHERE Email ='" + registration.Email + "'";
                    using (SqlCommand command = new SqlCommand(sqlEmail, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            reader.Read();
                            DuplicateUserEmail = reader.GetString(0);
                        }
                    }
                }
                catch (Exception)
                {

                    
                }

                if (DuplicateUserName != null)
                {
                    
                    ErrorMsg = "The name is already in use, please choose another";
                    return RedirectToPage("/Account/Register");
                }
                else if (DuplicateUserEmail != null)
                {
                    ErrorMsg = "The email is already in use, please choose another";
                    return RedirectToPage("/Account/Register"); ;
                }

                //Cria uma nova userList
                string sqlUserUpdate = "INSERT INTO UserList" +
                    "(UserName, Email, Password) VALUES" +
                    "(@UserName, @Email, @Password);";

                using (SqlCommand command = new SqlCommand(sqlUserUpdate, connection))
                {

                    command.Parameters.AddWithValue("@UserName", registration.UserName);
                    command.Parameters.AddWithValue("@Email", registration.Email);
                    command.Parameters.AddWithValue("@Password", registration.Password);

                    command.ExecuteNonQuery();

                }
            }
            return RedirectToPage("/Account/Login");
        }

        //public async Task SendVerificationEmail(string email)
        //{
        //    await _emailVerificationService.SendVerificationEmail(email, HttpContext);
        //}


    }
    public class Registration
    {
        [Required(ErrorMessage = "Please enter a username")]
        [Display(Name = "User Name")]
        public string UserName { get; set; }

        [Required(ErrorMessage = "Please enter your email")]
        [Display(Name = "Email Address")]
        [EmailAddress(ErrorMessage = "Please enter a valid email")]
        public string Email { get; set; }

        [Required(ErrorMessage = "Please enter a password")]
        [DataType(DataType.Password)]
        [Compare("ConfirmPassword", ErrorMessage = "Password does not match")]
        public string Password { get; set; }

        [Required(ErrorMessage = "Please confirm your password")]
        [Display(Name = "Confirm Password")]
        [DataType(DataType.Password)]
        public string ConfirmPassword { get; set; }


        public string Country { get; set; }
    }



    
}
