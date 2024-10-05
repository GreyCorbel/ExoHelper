using System;
using System.Net;
namespace ExoHelper
{
    public static class StringExtensions
    {
        //converts Exo size string to long
        public static long FromExoSize(this string input)
        {
            var start = input.IndexOf('(');
            var end = input.IndexOf(' ', start);
            long output = -1;
            long.TryParse(input.Substring(start + 1, end - start - 1).Replace(",", string.Empty), out output);
            return output;
        }
    }
}