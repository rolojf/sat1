#!/usr/bin/env nu

def otro [cualArchivo: string]     {
      let archivo = (open $cualArchivo)
      let tipo = ($archivo | get content.0.tag)
      if ($tipo == 'Emisor') {
      def obtnMonto [] {
           $archivo | get content.content.3.content.0.attributes.0.Base
      }

      def obtnIva [] {
           $archivo | get content.attributes.3.TotalImpuestosTrasladados
      }

      def obtnRFC [] {
           $archivo | get content.attributes.0.Rfc
      }

      def obtnFormaPago [] {
           $archivo | get attributes.FormaPago
      }

      def obtnSubTotal [] {
           $archivo | get attributes.SubTotal
      }

      def obtnNombre [] {
           $archivo | get content.0.attributes.Nombre
      }

       [[xml, nombre, rfc, pago, subTotal, monto, iva];[($cualArchivo), (obtnNombre), (obtnRFC), (obtnFormaPago), (obtnSubTotal), (obtnMonto), (obtnIva)]]
       } else {

      def obtnMonto [] {
           $archivo | get content.content.4.content.0.attributes.0.Base
      }

      def obtnIva [] {
           $archivo | get content.attributes.4.TotalImpuestosTrasladados
      }

      def obtnRFC [] {
           $archivo | get content.attributes.1.Rfc
      }

      def obtnFormaPago [] {
           $archivo | get attributes.FormaPago
      }

      def obtnSubTotal [] {
           $archivo | get attributes.SubTotal
      }

      def obtnNombre [] {
           $archivo | get content.1.attributes.Nombre
      }

       [[xml, nombre, rfc, pago, subTotal, monto, iva];[($cualArchivo), (obtnNombre), (obtnRFC), (obtnFormaPago), (obtnSubTotal), (obtnMonto), (obtnIva)]]
       }
}

def main [] {
    ls `**/*.xml` | get name | reduce -f [[xml, nombre, rfc, pago, subTotal, monto, iva];["cero", "ninguno", "xxx-xxx-xxx", 001, 0, 0, 0]] {|it, acc| $acc | append (try {(otro $it)} catch {[[xml, nombre, rfc, pago, subTotal, monto, iva];[($it), "error-nom", "error-rfc", 000, 0,0,0,]]})} | skip 1 | to html | save --force resumenDeXmls.html
    }
