/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.bean;

import java.util.List;

public class PendoladoVersion
{
    public static final String FICHAS_PENDOLADO = "/fichas-pendolado/";

    // Identificador del proyecto al que pertenece la versión.
    private Integer idProject;

    // Número de versión.
    private Integer numVersion;

    // Lista auxiliar para el .zul
    private List<PendoladoRevision> modelList;

    public PendoladoVersion(Integer idProject, Integer numVersion)
    {
        super();
        this.idProject = idProject;
        this.numVersion = numVersion;
    }

    public List<PendoladoRevision> getModelList()
    {
        return modelList;
    }

    public void setModelList(List<PendoladoRevision> pendoladoRevList)
    {
        this.modelList = pendoladoRevList;
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

    public String getFolderPath()
    {

        String basePath = System.getenv("SIRECA_HOME") + "/projects/";

        return basePath + idProject + FICHAS_PENDOLADO + "/" + numVersion + "/";

    }
}
