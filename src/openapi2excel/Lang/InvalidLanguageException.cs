namespace openapi2excel.core.Lang;

public class InvalidLanguageException : Exception
{
   public InvalidLanguageException() : base()
   {
   }

   public InvalidLanguageException(string message) : base(message)
   {
   }

   public InvalidLanguageException(string message, Exception innerException) : base(message, innerException)
   {
   }
}