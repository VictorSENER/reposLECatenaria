/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.bean;

import java.util.List;

import com.sener.sireca.web.util.IsJUnit;

public class ReplanteoVersion
{
    public static final String CALCULO_REPLANTEO = "/calculo-replanteo/";

    // Identificador del proyecto al que pertenece la versi�n.
    private Integer idProject;

    // N�mero de versi�n.
    private int numVersion;

    // Lista auxiliar para el .zul
    // private ArrayList<ModelReplanteoGrid> modelList;
    private List<ReplanteoRevision> modelList;

    public ReplanteoVersion(Integer idProject, Integer numVersion)
    {
        super();
        this.idProject = idProject;
        this.numVersion = numVersion;
    }

    public List<ReplanteoRevision> getModelList()
    {
        return modelList;
    }

    public void setModelList(List<ReplanteoRevision> replanteoRevList)
    {
        this.modelList = replanteoRevList;
    }

    public Integer getIdProject()
    {
        return idProject;
    }

    public void setIdProject(Integer idProject)
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

        String basePath = System.getenv("SIRECA_HOME");

        if (!IsJUnit.isJunitRunning())
            basePath += "/projects/";
        else
            basePath += "/projectTest/";

        return basePath + idProject + CALCULO_REPLANTEO + "/" + numVersion
                + "/";

    }
}
