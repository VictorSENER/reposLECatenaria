/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.worker;

import com.sener.sireca.web.bean.MontajeRevision;
import com.sener.sireca.web.service.MontajeServiceImpl;

public class MontajeWorker extends Thread
{
    // Revisión de la cual calcular el cuaderno de replanteo
    private MontajeRevision revision;
    private String catenaria;
    private double pkIni;
    private double pkFin;
    private boolean pdf;
    private boolean cad;

    public MontajeWorker(MontajeRevision revision, double pkIni, double pkFin,
            String catenaria, boolean pdf, boolean cad)
    {
        super();
        this.revision = revision;
        this.catenaria = catenaria;
        this.pkIni = pkIni;
        this.pkFin = pkFin;
        this.pdf = pdf;
        this.cad = cad;
    }

    @Override
    public void run()
    {
        MontajeServiceImpl service = new MontajeServiceImpl();
        service.calculateRevision(this.revision, this.pkIni, this.pkFin,
                this.catenaria, this.pdf, this.cad);
    }
}
