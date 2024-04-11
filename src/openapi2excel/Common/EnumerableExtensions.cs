namespace openapi2excel.core.Common
{
   internal static class EnumerableExtensions
   {
      public static void ForEach<T>(this IEnumerable<T> enumerable, Action<T> action)
      {
         foreach (var element in enumerable)
         {
            action(element);
         }
      }
   }
}