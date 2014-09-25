/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.util.List;

import com.sener.sireca.web.bean.Catenaria;

public interface CatenariaService
{
    public List<Catenaria> getAllCatenarias();

    public Catenaria getCatenariaById(int id);

    public Catenaria getCatenariaByTitle(String title);

    public List<String> getListCatenarias();
}
