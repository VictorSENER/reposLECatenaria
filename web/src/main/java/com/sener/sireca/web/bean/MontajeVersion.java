/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.bean;

import java.util.List;

public class MontajeVersion
{
    public static final String FICHAS_MONTAJE = "/fichas-montaje/";

    // Identificador del proyecto al que pertenece la versión.
    private Integer idProject;

    // Número de versión.
    private Integer numVersion;

    // Lista auxiliar para el .zul
    private List<MontajeRevision> modelList;

    public MontajeVersion(Integer idProject, Integer numVersion)
    {
        super();
        this.idProject = idProject;
        this.numVersion = numVersion;
    }

    public List<MontajeRevision> getModelList()
    {
        return modelList;
    }

    public void setModelList(List<MontajeRevision> montajeRevList)
    {
        this.modelList = montajeRevList;
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

        return basePath + idProject + FICHAS_MONTAJE + "/" + numVersion + "/";

    }
}
