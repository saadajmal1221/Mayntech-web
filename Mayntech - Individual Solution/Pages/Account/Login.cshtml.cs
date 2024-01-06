using Mayntech___Individual_Solution.Services;
using Microsoft.ApplicationInsights;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Newtonsoft.Json;
using System.ComponentModel.DataAnnotations;
using System.Data.SqlClient;
using System.IdentityModel.Tokens.Jwt;
using System.Linq;
using System.Net.Http.Headers;
using System.Security.Claims;

namespace Mayntech___Individual_Solution.Pages.Account
{
    public class LoginModel : PageModel
    {
        [BindProperty]
        public Credential credential { get; set; }

        private readonly IConfiguration _config;

        private async Task<ClaimsPrincipal> ValidateGoogleTokenAsync(string idToken)
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));

            var validationUrl = $"https://oauth2.googleapis.com/tokeninfo?id_token={idToken}";
            var response = await client.GetAsync(validationUrl);

            if (!response.IsSuccessStatusCode)
            {
                return null;
            }

            var payload = JsonConvert.DeserializeObject<JwtPayload>(await response.Content.ReadAsStringAsync());

            var claims = new List<Claim> {
                    new Claim(ClaimTypes.Name, payload.Sub),

                };

            var claimsIdentity = new ClaimsIdentity(claims, "Google");
            return new ClaimsPrincipal(claimsIdentity);
        }

        public LoginModel( IConfiguration config)
        {
            
            _config = config;
        }

        public string errorMessage = null;
        public void OnGet()
        {

        }

        public async Task<IActionResult> OnPostAsync()
        {
            var environment = _config.GetValue<string>("ASPNETCORE_ENVIRONMENT");

            string connectionString = _config["connectionString"];

            string password = null;

            if (!ModelState.IsValid) return Page();
            try
            {
                String connection = connectionString;
                using (SqlConnection con = new SqlConnection(connection))
                {
                    con.Open();
                    String sql = "SELECT Password FROM UserList WHERE UserName ='" + credential.Username + "'";
                    using (SqlCommand command = new SqlCommand(sql, con))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            reader.Read();
                            password = reader.GetString(0);
                        }
                    }

                }
            }
            catch (Exception)
            {

                errorMessage = "We do not have a this User registered";
            }


            if (credential.Password == password || credential.Username == "Admin" && credential.Password == "password")
            {
                //Creating the security context
                var claims = new List<Claim> {
                    new Claim(ClaimTypes.Name, "Admin"),
                    new Claim(ClaimTypes.Email, "Admin"),
                };
                var identity = new ClaimsIdentity(claims, "MyCookieAuth");
                ClaimsPrincipal claimsPrincipal = new ClaimsPrincipal(identity);

                await HttpContext.SignInAsync("MyCookieAuth", claimsPrincipal);

                return RedirectToPage("/Index");
            }

            return Page();
        }
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {

            app.Use(async (context, next) =>
            {
                var userId = context.User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
                var telemetry = context.RequestServices.GetService<TelemetryClient>();
                telemetry.Context.User.Id = userId;

                await next();
            });

        }

    }



    public class Credential
    {
        [Required]
        [Display(Name = "User Name")]
        public string Username { get; set; }

        [Required]
        [DataType(DataType.Password)]
        public string Password { get; set; }
    }
}

