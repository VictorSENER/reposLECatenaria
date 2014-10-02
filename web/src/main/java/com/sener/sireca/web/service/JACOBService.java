/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.util.List;

import com.jacob.com.Variant;

public interface JACOBService
{

    public boolean executeCoreCommand(String path, String fase,
            List<Variant> parameters);

}
