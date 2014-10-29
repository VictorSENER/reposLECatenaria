/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.page;

import java.util.ArrayList;

import javax.servlet.http.HttpSession;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.Sessions;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zul.Div;
import org.zkoss.zul.Grid;
import org.zkoss.zul.Html;

import com.sener.sireca.web.bean.DibujoRevision;
import com.sener.sireca.web.bean.DibujoVersion;
import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.service.ActiveProjectService;
import com.sener.sireca.web.service.DibujoService;
import com.sener.sireca.web.service.ProjectService;
import com.sener.sireca.web.util.SpringApplicationContext;

public class DibujoNotesPage extends SelectorComposer<Component>
{
    private static final long serialVersionUID = 1L;

    int numVersion;

    int numRevision;

    // Dialog components
    @Wire
    Grid notesList;

    @Wire
    Div notesContent;

    // Session data
    HttpSession session = (HttpSession) Sessions.getCurrent().getNativeSession();

    // Services
    ActiveProjectService actProj = (ActiveProjectService) SpringApplicationContext.getBean("actProj");
    DibujoService dibujoService = (DibujoService) SpringApplicationContext.getBean("dibujoService");
    ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");

    @Override
    public void doAfterCompose(Component comp) throws Exception
    {
        super.doAfterCompose(comp);

        numVersion = (Integer) Executions.getCurrent().getAttribute(
                "numVersion");
        numRevision = (Integer) Executions.getCurrent().getAttribute(
                "numRevision");
        Project project = projectService.getProjectById(actProj.getIdActive(session));
        DibujoVersion version = dibujoService.getVersion(project, numVersion);
        DibujoRevision revision = dibujoService.getRevision(version,
                numRevision);

        ArrayList<String> notes = dibujoService.getNotes(revision);

        String show = "";

        Html html;
        html = new Html();

        for (int i = 0; i < notes.size(); i++)
            if (i != 0)
                show += "<br>" + notes.get(i);
            else
                show += notes.get(i);

        html.setContent(show);

        html.setParent(notesContent);

    }

    @Listen("onClick = #goBack")
    public void doBackClick()
    {
        // Go back
        Executions.getCurrent().sendRedirect("/drawing");
    }

}
