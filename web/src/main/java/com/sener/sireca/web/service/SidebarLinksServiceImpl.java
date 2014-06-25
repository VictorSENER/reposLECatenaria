/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;

public class SidebarLinksServiceImpl implements SidebarLinksService
{

    HashMap<String, SidebarLink> linkMap = new LinkedHashMap<String, SidebarLink>();

    public SidebarLinksServiceImpl()
    {
        linkMap.put("user",
                new SidebarLink("user", "Usuarios", "/img/user.png", "/user"));
        linkMap.put(
                "catenary",
                new SidebarLink("catenary", "Configuraciones de Catenaria", "/img/catenary.png", "/catenary"));
        linkMap.put(
                "project",
                new SidebarLink("project", "Proyectos", "/img/project.png", "/project"));
        linkMap.put(
                "replanteo",
                new SidebarLink("replanteo", "Cuadernos de Replanteo", "/img/replanteo.png", "/replanteo"));
        linkMap.put(
                "drawing",
                new SidebarLink("drawing", "Planos de Replanteo", "/img/drawing.png", "/drawing"));
        linkMap.put(
                "pendolado",
                new SidebarLink("pendolado", "Fichas de Pendolado", "/img/ficha-pendolado.png", "/pendolado"));
        linkMap.put(
                "montaje",
                new SidebarLink("montaje", "Fichas de Montaje", "/img/ficha-montaje.png", "/montaje"));
    }

    public List<SidebarLink> getLinks()
    {
        return new ArrayList<SidebarLink>(linkMap.values());
    }

    public SidebarLink getLink(String name)
    {
        return linkMap.get(name);
    }

}