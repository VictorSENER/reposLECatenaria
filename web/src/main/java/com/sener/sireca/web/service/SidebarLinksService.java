/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */
package com.sener.sireca.web.service;

import java.util.List;

public interface SidebarLinksService
{
    public List<SidebarLink> getLinks();

    public SidebarLink getLink(String name);
}