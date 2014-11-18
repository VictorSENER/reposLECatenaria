/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.bean;

import java.util.List;

import com.sener.sireca.web.util.IsJUnit;

public class ReplanteoVersion
{
    public static final String CALCULO_REPLANTEO = "/calculo-replanteo/";

    // Identificador del proyecto al que pertenece la versión.
    private Integer idProject;

    // Número de versión.
    private int numVersion;

    // Indica si es eliminable
    private boolean canDelete;

    // Lista auxiliar para el .zul
    // private ArrayList<ModelReplanteoGrid> modelList;
    private List<ReplanteoRevision> modelList;

    public ReplanteoVersion(Integer idProject, int numVersion, boolean canDelete)
    {
        super();
        this.idProject = idProject;
        this.numVersion = numVersion;
        this.canDelete = canDelete;
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

    public boolean getCanDelete()
    {
        return canDelete;
    }

    public void setCanDelete(boolean canDelete)
    {
        this.canDelete = canDelete;
    }

    public String getFolderPath()
    {

        String basePath = System.getenv("SIRECA_HOME");

        if (!IsJUnit.isJunitRunning())
            basePath += "/projects/";
        else
            basePath += "/projectTest/";

        return basePath + idProject + CALCULO_REPLANTEO + numVersion + "/";

    }
}
