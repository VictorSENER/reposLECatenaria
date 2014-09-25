/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.controller;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.service.ActiveProjectService;
import com.sener.sireca.web.service.AuthenticationService;
import com.sener.sireca.web.service.DibujoService;
import com.sener.sireca.web.service.ProjectService;
import com.sener.sireca.web.service.ReplanteoService;
import com.sener.sireca.web.session.UserCredential;
import com.sener.sireca.web.util.SpringApplicationContext;

@Controller
public class AuthenticationController
{
    @Autowired
    AuthenticationService authService;
    @Autowired
    ActiveProjectService actProj;

    @RequestMapping(value = { "", "/", "index" }, method = RequestMethod.GET)
    public String index(ModelMap model, HttpServletRequest request,
            HttpSession session)
    {
        UserCredential cre = authService.getUserCredential(session);
        if (cre != null)
            return "redirect:main";
        else
            return "redirect:login";
    }

    @RequestMapping(value = "login", method = RequestMethod.GET)
    public String login(Model model)
    {
        return "login.zul";
    }

    @RequestMapping(value = "logout", method = RequestMethod.GET)
    public String logout(ModelMap model, HttpServletRequest request,
            HttpSession session)
    {
        authService.logout(session);
        actProj.setInactive(session);
        return "redirect:index";
    }

    @RequestMapping(value = "main", method = RequestMethod.GET)
    public String main(Model model)
    {
        return "main.zul";
    }

    @RequestMapping(value = "user", method = RequestMethod.GET)
    public String user(Model model)
    {
        return "user.zul";
    }

    @RequestMapping(value = "catenary", method = RequestMethod.GET)
    public String catenary(Model model)
    {
        return "catenary.zul";
    }

    @RequestMapping(value = "project", method = RequestMethod.GET)
    public String project(Model model)
    {
        return "project.zul";
    }

    @RequestMapping(value = "project/edit/{id}", method = RequestMethod.GET)
    public String editProject(@PathVariable Integer id, Model model)
    {
        return "projectEdit.zul";
    }

    private boolean isThereAnActiveProject(HttpSession session)
    {

        if (actProj.getIdActive(session) != 0)
            return true;
        return false;
    }

    @RequestMapping(value = "replanteo{action}", method = RequestMethod.GET)
    public String replanteo(@PathVariable String action, Model model,
            HttpServletRequest request, HttpSession session)
    {

        if (isThereAnActiveProject(session))
            return "replanteo.zul";

        return "nonActiveProject.zul";
    }

    @RequestMapping(value = "replanteo/new", method = RequestMethod.GET)
    public String newReplanteo(Model model)
    {
        return "replanteoNew.zul";
    }

    @RequestMapping(value = "replanteo/{action}/{numVersion}/{numRevision}", method = RequestMethod.GET)
    public String deleteShowRep(@PathVariable Integer numVersion,
            @PathVariable Integer numRevision, @PathVariable String action,
            Model model, HttpServletResponse response, HttpSession session)
    {
        if (isThereAnActiveProject(session))
        {
            if (action.equals("download"))
            {

                ActiveProjectService actProj = (ActiveProjectService) SpringApplicationContext.getBean("actProj");
                ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");
                ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");

                Project project = projectService.getProjectById(actProj.getIdActive(session));

                if (replanteoService.getRevision(
                        replanteoService.getVersion(project, numVersion),
                        numRevision).getCalculated())
                {
                    try
                    {

                        String path = replanteoService.getRevision(
                                replanteoService.getVersion(project, numVersion),
                                numRevision).getExcelPath();

                        String fileName = replanteoService.getRevision(
                                replanteoService.getVersion(project, numVersion),
                                numRevision).getExcelName();

                        // Get file as InputStream
                        InputStream is = new FileInputStream(path);

                        // Set file name
                        response.setHeader("Content-Disposition", "filename="
                                + fileName);

                        // Copy it to response's OutputStream
                        org.apache.commons.io.IOUtils.copy(is,
                                response.getOutputStream());
                        response.flushBuffer();
                    }
                    catch (IOException ex)
                    {
                        throw new RuntimeException("IOError writing file to output stream");
                    }
                }

            }
            else if (action.equals("progress"))
                return "replanteoProgress.zul";

            return "replanteo.zul";
        }
        return "nonActiveProject.zul";

    }

    @RequestMapping(value = "drawing{action}", method = RequestMethod.GET)
    public String drawing(Model model, HttpServletRequest request,
            HttpSession session)
    {
        if (isThereAnActiveProject(session))
            return "drawing.zul";
        return "nonActiveProject.zul";
    }

    @RequestMapping(value = "drawing/new", method = RequestMethod.GET)
    public String newDibujo(Model model)
    {
        return "drawingNew.zul";
    }

    @RequestMapping(value = "drawing/{action}/{numVersion}/{numRevision}", method = RequestMethod.GET)
    public String deleteShowDib(@PathVariable Integer numVersion,
            @PathVariable Integer numRevision, @PathVariable String action,
            Model model, HttpServletResponse response, HttpSession session)
    {
        if (isThereAnActiveProject(session))
        {
            if (action.equals("download"))
            {

                ActiveProjectService actProj = (ActiveProjectService) SpringApplicationContext.getBean("actProj");
                DibujoService dibujoService = (DibujoService) SpringApplicationContext.getBean("dibujoService");
                ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");

                Project project = projectService.getProjectById(actProj.getIdActive(session));

                if (dibujoService.getRevision(
                        dibujoService.getVersion(project, numVersion),
                        numRevision).getCalculated())
                {
                    try
                    {

                        String path = dibujoService.getRevision(
                                dibujoService.getVersion(project, numVersion),
                                numRevision).getAutoCadPath();

                        String fileName = dibujoService.getRevision(
                                dibujoService.getVersion(project, numVersion),
                                numRevision).getAutoCadName();

                        // Get file as InputStream
                        InputStream is = new FileInputStream(path);

                        // Set file name
                        response.setHeader("Content-Disposition", "filename="
                                + fileName);

                        // Copy it to response's OutputStream
                        org.apache.commons.io.IOUtils.copy(is,
                                response.getOutputStream());
                        response.flushBuffer();
                    }
                    catch (IOException ex)
                    {
                        throw new RuntimeException("IOError writing file to output stream");
                    }
                }

            }
            else if (action.equals("progress"))
                return "replanteoProgress.zul";

            return "drawing.zul";
        }
        return "nonActiveProject.zul";

    }

    @RequestMapping(value = "pendolado", method = RequestMethod.GET)
    public String pendolado(Model mode, HttpServletRequest request,
            HttpSession session)
    {
        if (isThereAnActiveProject(session))
            return "pendolado.zul";
        return "nonActiveProject.zul";
    }

    @RequestMapping(value = "montaje", method = RequestMethod.GET)
    public String montaje(Model model, HttpServletRequest request,
            HttpSession session)
    {
        if (isThereAnActiveProject(session))
            return "montaje.zul";
        return "nonActiveProject.zul";
    }
}
