namespace openapi2excel.core.Common;

internal class RowPointer(int rowNumber)
{
   public static implicit operator int(RowPointer d) => d.Get();
   public static explicit operator RowPointer(int b) => new(b);

   public void MoveNext(int rowCount = 1)
   {
      rowNumber += rowCount;
   }

   public void MovePrev(int rowCount = 1)
   {
      rowNumber -= rowCount;
   }

   public int Get() => rowNumber;

   public int GoTo(int row)
   {
      return rowNumber = row;
   }

   public RowPointer Copy() => new(rowNumber);
}