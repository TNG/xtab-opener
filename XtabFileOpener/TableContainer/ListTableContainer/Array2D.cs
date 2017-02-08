using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XtabFileOpener.TableContainer.ListTableContainer
{
    /// <summary>
    /// contains static methods that are useful for handling with two dimensional arrays
    /// </summary>
    public static class Array2D
    {
        internal static T[] getRow<T>(T[,] array, int rowNum)
        {
            T[] row = new T[array.GetLength(1)];
            int col = array.GetLowerBound(1);
            for (int i = 0; i < row.Length; i++)
            {
                row[i] = array[rowNum, col];
                col++;
            }
            return row;
        }

        internal static T[,] removeRow<T>(T[,] array, int rowNum)
        {
            if (array.GetLength(0) == 0)
                return array;

            T[,] newArray = (T[,])Array.CreateInstance(typeof(T),
                new int[2] { array.GetLength(0) - 1, array.GetLength(1) },
                new int[2] { array.GetLowerBound(0), array.GetLowerBound(1) });
            int col = 0;
            for (int i = newArray.GetLowerBound(0); i <= newArray.GetUpperBound(0); i++)
            {
                if (col == rowNum) col++;

                for (int j = newArray.GetLowerBound(0); j <= newArray.GetUpperBound(1); j++)
                    newArray[i, j] = array[col, j];

                col++;
            }
            return newArray;
        }

        /// <summary>
        /// creates a new array from an exisiting array with new dimensions
        /// </summary>
        /// <param name="array">old array</param>
        /// <param name="height">new height</param>
        /// <param name="width">new width</param>
        /// <returns>new array</returns>
        public static T[,] changeArraySize<T>(T[,] array, int height, int width)
        {
            int maxHeight = Math.Max(height, array.GetLength(0));
            int minHeight = Math.Min(height, array.GetLength(0));
            int maxWidth = Math.Max(width, array.GetLength(1));
            int minWidth = Math.Min(width, array.GetLength(1));

            int low0 = array.GetLowerBound(0);
            int low1 = array.GetLowerBound(1);

            T[,] newArray = (T[,])Array.CreateInstance(typeof(T),
                new int[2] { maxHeight, maxWidth },
                new int[2] { low0, low1 });
            //new string[maxHeight, maxWidth];

            for (int i = low0; i < low0 + minHeight; i++)
                for (int j = low1; j < low1 + minWidth; j++)
                    newArray[i, j] = array[i, j];
            return newArray;
        }

        /// <summary>
        /// creates a new array from an existing array, that has one more row
        /// </summary>
        /// <param name="array">old array</param>
        /// <param name="row">row that should be added</param>
        /// <returns>new array</returns>
        public static T[,] addRowToArray<T>(T[,] array, T[] row)
        {
            int indexOfNewRow = array.GetUpperBound(0) + 1;
            bool expandWidth = row.Length > array.GetLength(1);
            T[,] newArray = changeArraySize(array, array.GetLength(0) + 1, expandWidth ? row.Length : array.GetLength(1));
            int low1 = newArray.GetLowerBound(1);
            for (int i = row.GetLowerBound(0); i <= row.GetUpperBound(0); i++)
            {
                newArray[indexOfNewRow, low1] = row[i];
                low1++;
            }
            return newArray;
        }

        /// <summary>
        /// creates a new array of two exisiting arrays 
        /// </summary>
        /// <typeparam name="T">Type of the arrays</typeparam>
        /// <param name="arr1"></param>
        /// <param name="arr2"></param>
        /// <returns>array containing both arrays</returns>
        public static T[,] addArrayToArray<T>(T[,] arr1, T[,] arr2)
        {
            int newHeight = arr1.GetLength(0) + arr2.GetLength(0);
            int newWidth = Math.Max(arr1.GetLength(1), arr2.GetLength(1));
            T[,] newArray = changeArraySize(arr1, newHeight, newWidth);
            int row = newArray.GetLowerBound(0) + arr1.GetLength(0);
            for (int i = arr2.GetLowerBound(0); i <= arr2.GetUpperBound(0); i++)
            {
                int col = newArray.GetLowerBound(1);
                for (int j = arr2.GetLowerBound(1); j <= arr2.GetUpperBound(1); j++)
                {
                    newArray[row, col] = arr2[i, j];
                    col++;
                }
                row++;
            }
            return newArray;
        }

        public static bool areEqual(object[,] arr1, object[,] arr2)
        {
            int length0 = arr1.GetLength(0);
            int length1 = arr1.GetLength(1);
            if (length0 == arr2.GetLength(0) && length1 == arr2.GetLength(1))
            {
                int arr1Count0 = arr1.GetLowerBound(0);
                int arr2Count0 = arr2.GetLowerBound(0);
                for (int i = 0; i < length0; i++)
                {
                    int arr1Count1 = arr1.GetLowerBound(1);
                    int arr2Count1 = arr2.GetLowerBound(1);
                    for (int j = 0; j < length1; j++)
                    {
                        if (arr1[arr1Count0, arr1Count1] == null ^ arr2[arr2Count0, arr2Count1] == null)
                            return false;
                        if (arr1[arr1Count0, arr1Count1] == null && arr2[arr2Count0, arr2Count1] == null)
                            return true;
                        if (!arr1[arr1Count0, arr1Count1].Equals(arr2[arr2Count0, arr2Count1]))
                            return false;
                        arr1Count1++;
                        arr2Count1++;
                    }
                    arr1Count0++;
                    arr2Count0++;
                }
                return true;
            }
            else return false;
        }
    }
}
