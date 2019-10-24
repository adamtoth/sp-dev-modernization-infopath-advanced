﻿namespace SharePointPnP.Modernization.Framework.KQL
{
    /// <summary>
    /// Operators on KQL property filters
    /// </summary>
    public enum KQLPropertyOperator
    {
        Matches = 0,
        EqualTo = 1,
        LesserThan = 2,
        GreaterThan = 3,
        LesserThanOrEqualTo = 4,
        GreaterThanOrEqualTo = 5,
        DoesNoEqual = 6,
        Restriction = 7,
        Undefined = 100
    }
}
