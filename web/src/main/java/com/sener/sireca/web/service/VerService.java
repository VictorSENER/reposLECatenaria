/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.util.ArrayList;

public interface VerService
{
    public ArrayList<Integer> getVersions(String ruta);

    public boolean getVersion(String ruta, int version);

    public int getLastVersion(String ruta);

}
