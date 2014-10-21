/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.session;

import java.io.Serializable;

public class ActiveProject implements Serializable
{
    private static final long serialVersionUID = 1L;

    int idSelectedProject = 0;
    String selectedProject;

    public ActiveProject(int idProj, String titleProj)
    {
        this.idSelectedProject = idProj;
        this.selectedProject = titleProj;
    }

    public int getIdSelectedProject()
    {
        return idSelectedProject;
    }

    public void setIdSelectedProject(int idSelectedProject)
    {
        this.idSelectedProject = idSelectedProject;
    }

    public String getSelectedProject()
    {
        return selectedProject;
    }

    public void setSelectedProject(String selectedProject)
    {
        this.selectedProject = selectedProject;
    }

}
