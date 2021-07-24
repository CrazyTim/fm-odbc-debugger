using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class Sql
    {
        public static class Syntax
        {
            public static class Filemaker
            {
                public static readonly Regex String = new Regex(@"""""|''|"".*?[^\\]""|'.*?[^\\]'", RegexOptions.Compiled);
                public static readonly Regex Number = new Regex(@"\b\d+[\.]?\d*([eE]\-?\d+)?\b", RegexOptions.Compiled);
                public static readonly Regex Comment1 = new Regex("--.*$", RegexOptions.Multiline | RegexOptions.Compiled);
                public static readonly Regex Comment2 = new Regex(@"(/\*.*?\*/)|(/\*.*)", RegexOptions.Singleline | RegexOptions.Compiled);
                public static readonly Regex Comment3 = new Regex(@"(/\*.*?\*/)|(.*\*/)", RegexOptions.Singleline | RegexOptions.RightToLeft | RegexOptions.Compiled);
                public static readonly Regex Keywords = new Regex(@"\b(ALTER|CATALOG|CREATE|DELETE|DROP|EXEC|EXECUTE|FROM|IN|INSERT|INTO|OPTION|OUTPUT|SELECT|SET|TRUNCATE|WHERE|WITH|ADD|ARE|AS|ASC|AT|BY|CASE|COLUMN|COMMIT|CONNECT|CURRENT|DEFAULT|DESC|DISTINCT|ELSE|END|EXTERNAL|FETCH|FIRST|FOR|GLOBAL|GO|GOTO|GROUP|HAVING|INDEX|KEY|NO|OF|OFFSET|ON|ONLY|OPEN|ORDER|PERCENT|PRIMARY|ROW|ROWS|TABLE|THEN|TIES|TO|UNION|UNIQUE|UPDATE|VALUE|VALUES|WHEN)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
                public static readonly Regex Operators = new Regex(@"\b(AND|OR|INNER JOIN|LEFT OUTER JOIN|LIKE|NOT|IS|NULL|BETWEEN|IN|EXISTS|ANY|ALL)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
                public static readonly Regex SpecialKeywords = new Regex(@"\b(FileMaker_Tables|FileMaker_Fields|ROWMODID|ROWID)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
                public static readonly Regex Functions = new Regex(@"\b(ABS|ATAN|ATAN2|AVE|AVG|CAST|CEIL|CEILING|CHAR|CHR|CHARACTER_LENGTH|CHAR_LENGTH|COALESCE|CONVERT|COUNT|CURDATE|CURRENT_DATE|CURRENT_TIME|CURRENT_TIMESTAMP|CURRENT_USER|CURTIME|CURTIMESTAMP|DATE|DATEVAL|DAY|DAYNAME|DAYOFWEEK|DEG|DEGREES|EXP|FLOOR|GETAS|HOUR|INT|LEFT|LENGTH|LOWER|LN|LOG|LTRIM|MAX|MIN|MINUTE|MOD|MONTH|MONTHNAME|NULLIF|NUMVAL|PI|PAD|RADIANS|RIGHT|ROUND|RTRIM|SECOND|SIGN|SIN|SPACE|SQRT|STRVAL|SUBSTR|SUBSTRING|SUM|TAN|TIMESTAMPVAL|TIMEVAL|TODAY|TRIM|UPPER|USER|USERNAME|YEAR)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
                public static readonly Regex Types = new Regex(@"\b(NUMERIC|DECIMAL|INT|DATE|TIME|TIMESTAMP|VARCHAR|CHARACTER VARYING|BLOB|VARBINARY|LONGVARBINARY|BINARY VARYING)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

                /// <summary> These keywords don't do anything in FileMaker, but they still can't be used as identifiers. </summary>
                public static readonly Regex ReservedKeywords = new Regex(@"\b(BIT|BIT_LENGTH|BOOLEAN|DEC|FLOAT|INTEGER|NCHAR|REAL|SMALLINT|TIMEZONE_HOUR|TIMEZONE_MINUTE|WHENEVER|WORK|WRITE|ZONE|VIEW|USAGE|USING|UNKNOWN|TRAILING|TRANSACTION|TRANSLATE|TRANSLATION|TRUE|TEMPORARY|SCHEMA|SCROLL|SECTION|SESSION|SESSION_USER|SIZE|SOME|SQL|SQLCODE|SQLERROR|SQLSTATE|SYSTEM_USER|PRIOR|PRIVILEGES|PROCEDURE|PUBLIC|READ|REFERENCES|RELATIVE|RESTRICT|REVOKE|ROLLBACK|POSITION|PRECISION|PREPARE|PRESERVE|OUTER|OVERLAPS|PART|PARTIAL|OCTET_LENGTH|LANGUAGE|LAST|LEADING|LEVEL|LOCAL|MATCH|MODULE|NAMES|NATIONAL|NATURAL|NEXT|INDICATOR|INITIALLY|INPUT|INSENSITIVE|INTERSECT|INTERVAL|ISOLATION|IDENTITY|IMMEDIATE|GRANT|FOREIGN|FOUND|FULL|GET|EXTRACT|FALSE|ESCAPE|EVERY|EXCEPT|EXCEPTION|DOMAIN|DOUBLE|DESCRIBE|DESCRIPTOR|DIAGNOSTICS|DISCONNECT|DEFERRABLE|DEFERRED|DECLARE|DEALLOCATE|CURSOR|CORRESPONDING|CONTINUE|CONSTRAINT|CONSTRAINTS|CONNECTION|COLLATION|COLLATE|CLOSE|CHECK|CASCADED|CASCADE|BOTH|ASSERTION|AUTHORIZATION|ALLOCATE|ABSOLUTE|ACTION|BEGIN|END_EXEC|INNER|EXCEPTION|JOIN|CROSS)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
            }
        }
    }
}
