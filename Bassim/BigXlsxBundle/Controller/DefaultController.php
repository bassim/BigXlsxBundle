<?php

namespace Bassim\BigXlsxBundle\Controller;

use Symfony\Bundle\FrameworkBundle\Controller\Controller;

class DefaultController extends Controller
{
    public function indexAction($name)
    {
        return $this->render('BassimBigXlsxBundle:Default:index.html.twig', array('name' => $name));
    }
}
