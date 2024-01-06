namespace Mayntech___Individual_Solution.Services
{
    public interface IEmailSender
    {
        Task SendEmailAsync(string email, string subject, string message)
        {
            return Task.CompletedTask;
        }
    }
}
