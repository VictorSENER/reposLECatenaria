/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.bean;

import java.util.List;

import com.sener.sireca.web.util.IsJUnit;

public class DibujoVersion
{
    public static final String DIBUJO_REPLANTEO = "/dibujo-replanteo/";

    // Identificador del proyecto al que pertenece la versión.
    private Integer idProject;

    // Número de versión.
    private Integer numVersion;

    // Indica si es eliminable
    private boolean canDelete;

    // Lista auxiliar para el .zul
    private List<DibujoRevision> modelList;

    public DibujoVersion(Integer idProject, Integer numVersion,
            boolean canDelete)
    {
        super();
        this.idProject = idProject;
        this.numVersion = numVersion;
        this.canDelete = canDelete;
    }

    public List<DibujoRevision> getModelList()
    {
        return modelList;
    }

    public void setModelList(List<DibujoRevision> dibujoRevList)
    {
        this.modelList = dibujoRevList;
    }

    public Integer getIdProject()
    {
        return idProject;
    }

    public void setIdProject(Integer idProject)
    {
        this.idProject = idProject;
    }

    public Integer getNumVersion()
    {
        return numVersion;
    }

    public void setNumVersion(Integer numVersion)
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

        return basePath + idProject + DIBUJO_REPLANTEO + "/" + numVersion + "/";

    }
}
