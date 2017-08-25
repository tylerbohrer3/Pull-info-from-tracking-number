^w::
SetTitleMatchMode, 2
FormatTime, Year, , yyyy
TrackingNumber =                                                                ;Put your tracking number in here
        StringReplace, TrackingNumber, TrackingNumber, %A_Space%, , 1
        StringReplace, TrackingNumber, TrackingNumber, %A_Tab%, , 1
        StringReplace, TrackingNumber, TrackingNumber, `,, , 1
        StringReplace, TrackingNumber, TrackingNumber, `., , 1
        StringReplace, TrackingNumber, TrackingNumber, `?, , 1
        StringReplace, TrackingNumber, TrackingNumber, `-, , 1
        StringReplace, TrackingNumber, TrackingNumber, `*, , 1
        StringReplace, TrackingNumber, TrackingNumber, `+, , 1
        StringReplace, TrackingNumber, TrackingNumber, `r, , 1
        StringReplace, TrackingNumber, TrackingNumber, `n, , 1
        StringReplace, TrackingNumber, TrackingNumber, ``, , 1
        StringReplace, TrackingNumber, TrackingNumber, `@, , 1
        StringReplace, TrackingNumber, TrackingNumber, `#, , 1
        StringReplace, TrackingNumber, TrackingNumber, `&, , 1
        StringReplace, TrackingNumber, TrackingNumber, `_, , 1
        StringReplace, TrackingNumber, TrackingNumber, `', , 1
        StringReplace, TrackingNumber, TrackingNumber, `", , 1
        StringReplace, TrackingNumber, TrackingNumber, `;, , 1
        StringReplace, TrackingNumber, TrackingNumber, `:, , 1
        StringReplace, TrackingNumber, TrackingNumber, `(, , 1
        StringReplace, TrackingNumber, TrackingNumber, `), , 1
        StringReplace, TrackingNumber, TrackingNumber, `=, , 1
        SetFormat, float, 34
        FormatTime, Date,,M/d
        StringLen, Length, TrackingNumber
        wb := ComObjCreate("InternetExplorer.Application")
        wb.Visible := True
        If (Length = 12){
            StringSplit, characters, TrackingNumber
            newNum = 0
            Loop, 9{
                newNum += 1
                character = characters%newNum%
                If %character% is alpha
                    MsgBox, 4, Error`, Tracking number not recognized., `"%TrackingNumber%`" may not be a valid Tracking Number. Press `"Yes`" to check this Number on the carrier's website, or `"No`" to exit.
                    IfMsgBox Yes
                    {
                        run http://ltl.upsfreight.com/shipping/tracking/TrackingDetail.aspx?TrackProNumber=%TrackingNumber%
                        return
                    }
            }
            newNum += 1
            character = characters%newNum%
            If %character% is alpha                                     ;strip letters from UPS Freight numbers
            {
                newNum += 1
                character = characters%newNum%
                If %character% is alpha
                {
                    newNum += 1
                    character = characters%newNum%
                    If %character% is alpha
                    {
                        TrackingNumber = %characters1%%characters2%%characters3%%characters4%%characters5%%characters6%%characters7%%characters8%%characters9%
                        Length = 9
                    }
                }
            }
        }
        DelDate = 
        if (Length = 9){
            webAddress = http://ltl.upsfreight.com/shipping/tracking/TrackingDetail.aspx?TrackProNumber=%TrackingNumber%
            wb.Navigate(WebAddress)
            IELoad(wb)
            If (wb.document.getElementByID("app_ctl00_lblDeliverStatus").InnerText = "Delivered"){    
                DelDate := wb.document.getElementByID("app_ctl00_lblDeliveredOn").InnerText                         ;Get date package was delivered
            }
            If (wb.document.getElementByID("app_ctl00_lblDeliverStatus").InnerText = "On Vehicle For Delivery"){    
                DelDate := wb.document.getElementByID("app_ctl00_lblScheduledDelivery").InnerText                 ;Get scheduled delivery date
            }
            If (wb.document.getElementByID("app_ctl00_lblDeliverStatus").InnerText = "In Transit"){    
                DelDate := wb.document.getElementByID("app_ctl00_lblScheduledDelivery").InnerText                  ;Get scheduled delivery date
            }
            If (DelDate = ""){
                MsgBox, 4, Error`, Tracking number not recognized., `"%TrackingNumber%`" may not be a valid Tracking newNumber. Press `"Yes`" to check this number on the carrier's website, or `"No`" to exit.
                IfMsgBox Yes
                    return
            }
        }
        If (Length = 10 OR Length = 12 OR Length = 15 OR Length = 22 OR Length = 34){
            webAddress = https://www.fedex.com/apps/fedextrack/?action=track&tracknumbers=%tracking%&action=track&language=english&state=0&cntry_code=us
            wb.Navigate(WebAddress)
            IELoad(wb)
            ComObjError(false)
            While (Not wb.document.all.trackNumbers.OuterHTML){
                sleep 20
            }
            ComObjError(true)
            loop % (divs := wb.Document.getElementsbytagname("div")).length{
                ComObjError(false)
                If (wb.Document.getElementsbytagname("div")[A_index+130].InnerText = "Scheduled delivery: ")
                    delivery = Scheduled
                If (wb.Document.getElementsbytagname("div")[A_index+130].InnerText = "Actual delivery: ")
                    delivery = Delivered
                variable := wb.Document.getElementsbytagname("div")[A_index+130].OuterHTML
                stringsplit, variable, variable
                If (variable6 = "c" AND variable13 = "s" AND variable32 = "d" AND variable35 = "e"){
                    If (delivery = "Scheduled"){
                        DelDate := wb.Document.getElementsbytagname("div")[A_index+130].InnerText
                        If (DelDate = "Pending"){
                            Shipdate := wb.Document.getElementsbytagname("div")[A_index+130].InnerText                          ;Get date package was shipped
                        }
                    }
                    else If (delivery = "Delivered"){
                        DelDate := wb.Document.getElementsbytagname("div")[A_index+130].InnerText                               ;Get date package was delivered
                        break
                    }
                }
            }
            If (DelDate = ""){
                MsgBox, 4, Error`, Tracking number not recognized., `"%TrackingNumber%`" may not be a valid tracking number. Press `"Yes`" to check this number on the carrier's website, or `"No`" to exit.
                IfMsgBox Yes
                    return
            }
        }
        delivery = 
        wb.quit()
        }
    }
}
Length =
return
