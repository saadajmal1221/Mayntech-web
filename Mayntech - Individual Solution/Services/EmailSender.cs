using Azure.Core;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Net.Mail;

namespace Mayntech___Individual_Solution.Services
{

    public class EmailVerificationService
    {
        private readonly IEmailSender _emailSender;
        private readonly UserManager<IdentityUser> _userManager;
        private readonly IUrlHelper _urlHelper;

        public EmailVerificationService(IEmailSender emailSender, UserManager<IdentityUser> userManager, IUrlHelper urlHelper)
        {
            _emailSender = emailSender;
            _userManager = userManager;
            _urlHelper = urlHelper;
        }

        public async Task SendVerificationEmail(string email, HttpContext httpContext)
        {
            var user = await _userManager.FindByEmailAsync(email);
            if (user == null)
            {
                throw new Exception($"Unable to find user with email: {email}");
            }

            var token = await _userManager.GenerateEmailConfirmationTokenAsync(user);
            var confirmationLink = _urlHelper.Action("ConfirmEmail", "Account",
                new { userId = user.Id, token = token }, httpContext.Request.Scheme);

            await _emailSender.SendEmailAsync(email, "Confirm your email",
                $"Please confirm your account by clicking this link: <a href='{confirmationLink}'>link</a>");
        }
    }

    public class UrlService
    {
        private readonly IHttpContextAccessor _httpContextAccessor;
        private readonly LinkGenerator _linkGenerator;

        public UrlService(IHttpContextAccessor httpContextAccessor, LinkGenerator linkGenerator)
        {
            _httpContextAccessor = httpContextAccessor;
            _linkGenerator = linkGenerator;
        }
        public string GenerateUrl(string routeName, object routeValues = null)
        {
            var httpContext = _httpContextAccessor.HttpContext;
            var url = _linkGenerator.GetPathByRouteValues(httpContext, routeName, routeValues);
            return url;
        }
    }

    public class EmailSender : IEmailSender
    {
        private readonly SmtpClient _smtpClient;

        public EmailSender(string host, int port)
        {
            _smtpClient = new SmtpClient(host, port);
        }

        public async Task SendEmailAsync(string email, string subject, string message)
        {
            var mailMessage = new MailMessage
            {
                From = new MailAddress("noreply@yourdomain.com"),
                Subject = subject,
                Body = message,
                IsBodyHtml = true
            };
            mailMessage.To.Add(new MailAddress(email));
            await _smtpClient.SendMailAsync(mailMessage);
        }
    }

}

