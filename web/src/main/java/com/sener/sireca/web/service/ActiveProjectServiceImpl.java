/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import javax.servlet.http.HttpSession;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.sener.sireca.web.session.ActiveProject;

@Service("actProj")
public class ActiveProjectServiceImpl implements ActiveProjectService
{

    @Autowired
    ProjectService projectService;

    public void setActive(HttpSession session, int idProj, String titleProj)
    {
        ActiveProject proj = new ActiveProject(idProj, titleProj);

        session.setAttribute("activeProject", proj);

    }

    public int getIdActive(HttpSession session)
    {
        ActiveProject proj = (ActiveProject) session.getAttribute("activeProject");

        if (proj == null)
            return 0;

        return proj.getIdSelectedProject();
    }

    // public boolean getActive(HttpSession session)
    // {
    // // TODO Auto-generated method stub
    // return null;
    // }

}
