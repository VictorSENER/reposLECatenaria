/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.dao;

import java.util.List;

import com.sener.sireca.web.bean.Catenaria;

public interface CatenariaDao
{

    public List<Catenaria> getAllCatenarias();

    public Catenaria getCatenariaById(int id);

}