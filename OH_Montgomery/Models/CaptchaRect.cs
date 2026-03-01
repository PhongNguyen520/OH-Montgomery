namespace OH_Montgomery.Models;

/// <summary>Bounding rect for captcha image extraction (used by DomHelper.ExtractCaptchaBase64Async).</summary>
public record CaptchaRect(double x, double y, double width, double height);
