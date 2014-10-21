/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.util;

public class IsJUnit
{
    private static boolean junitRunning = false;

    public static boolean isJunitRunning()
    {
        return junitRunning;
    }

    public static void setJunitRunning(boolean junitRunning)
    {
        IsJUnit.junitRunning = junitRunning;
    }

}
