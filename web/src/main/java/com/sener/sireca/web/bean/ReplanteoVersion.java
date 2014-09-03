/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.bean;

public class ReplanteoVersion
{
    // Identificador del proyecto al que pertenece la versión.
    private int idProject;

    // Número de versión.
    private int numVersion;

    public int getIdProject()
    {
        return idProject;
    }

    public void setIdProject(int idProject)
    {
        this.idProject = idProject;
    }

    public int getNumVersion()
    {
        return numVersion;
    }

    public void setNumVersion(int numVersion)
    {
        this.numVersion = numVersion;
    }

    public String getFolderPath()
    {
        String basePath = System.getenv("SIRECA_HOME") + "/" + "projects";
        return basePath + "/" + idProject + "/" + "replanteo" + "/"
                + numVersion;
    }
}
