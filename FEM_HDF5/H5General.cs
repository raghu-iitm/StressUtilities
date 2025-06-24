using System;
using System.Collections.Generic;
using System.Linq;

namespace Nastranh5
{
    public class H5General
    {

        public H5General() { }

        public List<Int64> GetEntityList(string entityList)
        {
            //string[] listInter;
            //string[] ExpandedList;
            //long start;
            //long end;
            long increment = 1;
            List<Int64> ListFinal = new List<Int64>();
            if (string.IsNullOrEmpty(entityList))
            {
                return new List<Int64>();
            }

            string[]  listInter = entityList.Split(new char[] { ';', ',', '\n', '\t', ' ' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string listItem in listInter)
            {
                if (listItem.Contains(":"))
                {
                    string[] ExpandedList = listItem.Split(new char[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
                    long start = Convert.ToInt64(ExpandedList[0]);
                    long end = Convert.ToInt64(ExpandedList[1]);

                    if (ExpandedList.Length >= 3)
                    {
                        increment = Convert.ToInt64(ExpandedList[2]);
                    }

                    for (long i = start; i <= end; i = i + increment)
                    {
                        ListFinal.Add(i);
                    }

                }
                else
                {
                    if (long.TryParse(listItem, out long entity))   //listItem.All(char.IsDigit);  
                    {
                        ListFinal.Add(Convert.ToInt64(listItem));
                    }
                }
            }

            return ListFinal.Distinct().OrderBy(x => x).ToList();
        }

        public string GetDatasetType(string dataset)
        {
            switch (dataset)
            {
                case "TRIA3":
                case "TRIA6":
                case "QUAD4":
                    return "SS2D";
                case "TRIA3_CPLX":
                case "TRIA6_CPLX":
                case "QUAD4_CPLX":
                    return "CPLXSS2D";
                case "HEXA":
                case "PENTA":
                case "TETRA":
                    return "SS3D";
                case "HEXA_CPLX":
                case "PENTA_CPLX":
                case "TETRA_CPLX":
                    return "CPLXSS3D";
                default:
                    return "NONE";
            }
        }

        /* public static object BinarySearch(int[] inputArray, int key, int min, int max, out int[] IndexDict)
          {
              //int min = 0;
              //int max = inputArray.Length - 1;
              while (min <= max)
              {
                  int mid = (min + max) / 2;
                  if (key == inputArray[mid])
                  {
                      //return ++mid;
                  }
                  else if (key < inputArray[mid])
                  {
                      max = mid - 1;
                  }
                  else
                  {
                      min = mid + 1;
                  }
              }
              return "Nil";
          }*/


    }
}
