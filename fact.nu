#!/usr/bin/env nu

def otro [cualArchivo: string]     {
      let archivo = (open $cualArchivo)


      def obtnMonto [] {
           $archivo | get content | filter {|x|$x.tag == "Impuestos" } | get content.0 | filter {|x|$x.tag == "Traslados" } | get content.0 | filter {|x|$x.tag == "Traslado" } | get attributes.Base.0
      }

      def obtnIva [] {
           $archivo | get content | filter {|x|$x.tag == "Impuestos" } | get content.0 | filter {|x|$x.tag == "Traslados" } | get content.0 | filter {|x|$x.tag == "Traslado" } | get attributes.Importe.0
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
           $archivo | get content.0.attributes.Nombre
      }

      def obtnRetIva [] {
           let extracto = $archivo | get content | filter {|x|$x.tag == "Impuestos" } | get content | flatten | filter {|x|$x.tag == "Retenciones" }
           if ($extracto | is-empty ) { 0 } else { $extracto | get content | flatten | get attributes | filter {|x|$x.Impuesto == "002" } |  get Importe.0 }
      }

      def obtnRetIsr [] {
           let extracto = $archivo | get content | filter {|x|$x.tag == "Impuestos" } | get content | flatten | filter {|x|$x.tag == "Retenciones" }
           if ( $extracto | is-empty ) { 0 } else { $extracto | get content | flatten | get attributes | filter {|x|$x.Impuesto == "001" } |  get Importe.0 }
      }

      [[xml, nombre, rfc, pago, subTotal, monto, iva, retIva, RetIsr]; [($cualArchivo), (obtnNombre), (obtnRFC), (obtnFormaPago), (obtnSubTotal), (obtnMonto), (obtnIva), (obtnRetIva), (obtnRetIsr)]]
}

def main [] {
    ls -s *.xml | get name | reduce -f [[xml, nombre, rfc, pago, subTotal, monto, iva, retIsr, retIva]; [uno, dos, tres, 4, 5, 6, 7, 8, 9] ] {|it, acc| $acc | append (otro $it) } | skip 1 | to html  | save --force resumenDeFacs.html
}
