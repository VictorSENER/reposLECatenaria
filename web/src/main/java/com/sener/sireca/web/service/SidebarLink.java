/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */
package com.sener.sireca.web.service;

import java.io.Serializable;

public class SidebarLink implements Serializable
{
    private static final long serialVersionUID = 1L;

    String name;
    String label;
    String iconUri;
    String url;

    public SidebarLink(String name, String label, String iconUri, String url)
    {
        this.name = name;
        this.label = label;
        this.iconUri = iconUri;
        this.url = url;
    }

    public String getName()
    {
        return name;
    }

    public String getLabel()
    {
        return label;
    }

    public String getIconUri()
    {
        return iconUri;
    }

    public String getUrl()
    {
        return url;
    }
}