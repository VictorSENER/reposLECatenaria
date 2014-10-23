/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.session;

import java.io.Serializable;

public class ActiveProject implements Serializable
{
    private static final long serialVersionUID = 1L;

    Integer idActiveProject = 0;
    String titleActiveProject = "";

    public ActiveProject(Integer idProj, String titleProj)
    {
        this.idActiveProject = idProj;
        this.titleActiveProject = titleProj;
    }

    public Integer getIdSelectedProject()
    {
        return idActiveProject;
    }

    public void setIdSelectedProject(Integer idSelectedProject)
    {
        this.idActiveProject = idSelectedProject;
    }

    public String getSelectedProject()
    {
        return titleActiveProject;
    }

    public void setSelectedProject(String selectedProject)
    {
        this.titleActiveProject = selectedProject;
    }

}
