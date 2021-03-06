/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.bean;

import java.util.List;

public class DibujoVersion
{
    public static final String DIBUJO_REPLANTEO = "/dibujo-replanteo/";

    // Identificador del proyecto al que pertenece la versi?n.
    private Integer idProject;

    // N?mero de versi?n.
    private Integer numVersion;

    // Lista auxiliar para el .zul
    private List<DibujoRevision> modelList;

    public DibujoVersion(Integer idProject, Integer numVersion)
    {
        super();
        this.idProject = idProject;
        this.numVersion = numVersion;
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

    public String getFolderPath()
    {

        String basePath = System.getenv("SIRECA_HOME") + "/projects/";

        return basePath + idProject + DIBUJO_REPLANTEO + "/" + numVersion + "/";

    }
}
