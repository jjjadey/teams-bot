// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {
    TurnContext,
    MessageFactory,
    TeamsActivityHandler,
    CardFactory,
    ActionTypes
} = require('botbuilder');
const { AdaptiveCards } = require("@microsoft/adaptivecards-tools");
// const { CardTemplate } = require("./CardTemplate.json")

class BotActivityHandler extends TeamsActivityHandler {
    constructor() {
        super();
    }

    /* Building a messaging extension search command is a two step process.
        (1) Define how the messaging extension will look and be invoked in the client.
            This can be done from the Configuration tab, or the Manifest Editor.
            Learn more: https://aka.ms/teams-me-design-search.
        (2) Define how the bot service will respond to incoming search commands.
            Learn more: https://aka.ms/teams-me-respond-search.
        
        NOTE:   Ensure the bot endpoint that services incoming messaging extension queries is
                registered with Bot Framework.
                Learn more: https://aka.ms/teams-register-bot. 
    */

    // Invoked when the service receives an incoming search query.
    async handleTeamsMessagingExtensionQuery(context, query) {
        const axios = require('axios');
        const querystring = require('querystring');
        const { commandId } = query;

        const searchQuery = query.parameters[0].value;
        const response = await axios.get(`http://registry.npmjs.com/-/v1/search?${querystring.stringify({ text: searchQuery, size: 8 })}`);

        const attachments = [];
        response.data.objects.forEach(obj => {

            const card = CardFactory.heroCard(obj.package.date); //never display but we required it to add preview invoke

            const preview = CardFactory.heroCard(obj.package.name, `version: ${obj.package.version}`); // list item
            const passData = { commandId, name: obj.package.name, description: obj.package.description, test: '111111111111' } //can be any obj
            preview.content.tap = { type: 'invoke', value: passData }; // required for trigger handleTeamsMessagingExtensionSelectItem

            const attachment = { ...card, preview };
            attachments.push(attachment);

        });

        return {
            composeExtension: {
                type: 'result',
                attachmentLayout: 'list',
                attachments: attachments
            }
        };
    }

    // Invoked when the user selects an item from the search result list returned above.
    async handleTeamsMessagingExtensionSelectItem(context, obj) {
        console.log('--------------selected item-------------------')
        console.log(obj);
        const { commandId, name, description, test } = obj;

        let card;
        if (commandId === 'heroCardSearch') {
            card = CardFactory.heroCard(description, test, [{ url: "https://adaptivecards.io/content/cats/3.png" }]);
        } else if (commandId === 'adaptiveCardSearch') {
            const template = {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard",
                "version": "1.6",
                "speak": "<s>Flight KL0605 to San Fransisco has been delayed.</s><s>It will not leave until 10:10 AM.</s>",
                "body": [
                    {
                        "type": "TextBlock",
                        "text": "Your Flight Update",
                        "wrap": true,
                        "style": "heading"
                    },
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "auto",
                                "items": [
                                    {
                                        "type": "Image",
                                        "size": "small",
                                        "url": "https://adaptivecards.io/content/airplane.png",
                                        "altText": "Airplane"
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Flight Status",
                                        "horizontalAlignment": "right",
                                        "isSubtle": true,
                                        "wrap": true
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": "DELAYED",
                                        "horizontalAlignment": "right",
                                        "spacing": "none",
                                        "size": "large",
                                        "color": "attention",
                                        "wrap": true
                                    }
                                ]
                            }
                        ]
                    },
                    {
                        "type": "ColumnSet",
                        "separator": true,
                        "spacing": "medium",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Passengers",
                                        "isSubtle": true,
                                        "weight": "bolder",
                                        "wrap": true
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": "${underName.name}",
                                        "spacing": "small",
                                        "wrap": true
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": "Jeremy Goldberg",
                                        "spacing": "small",
                                        "wrap": true
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": "Evan Litvak",
                                        "spacing": "small",
                                        "wrap": true
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": "auto",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Seat",
                                        "horizontalAlignment": "right",
                                        "isSubtle": true,
                                        "weight": "bolder",
                                        "wrap": true
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": "14A",
                                        "horizontalAlignment": "right",
                                        "spacing": "small",
                                        "wrap": true
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": "14B",
                                        "horizontalAlignment": "right",
                                        "spacing": "small",
                                        "wrap": true
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": "14C",
                                        "horizontalAlignment": "right",
                                        "spacing": "small",
                                        "wrap": true
                                    }
                                ]
                            }
                        ]
                    },
                    {
                        "type": "ColumnSet",
                        "spacing": "medium",
                        "separator": true,
                        "columns": [
                            {
                                "type": "Column",
                                "width": 1,
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Flight",
                                        "isSubtle": true,
                                        "weight": "bolder",
                                        "wrap": true
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": "${reservationFor.flightNumber}",
                                        "spacing": "small",
                                        "wrap": true
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": 1,
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Departs",
                                        "isSubtle": true,
                                        "horizontalAlignment": "center",
                                        "weight": "bolder",
                                        "wrap": true
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": "{{TIME(${string(reservationFor.departureTime)})}}",
                                        "color": "attention",
                                        "weight": "bolder",
                                        "horizontalAlignment": "center",
                                        "spacing": "small",
                                        "wrap": true
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": 1,
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Arrives",
                                        "isSubtle": true,
                                        "horizontalAlignment": "right",
                                        "weight": "bolder",
                                        "wrap": true
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": "{{TIME(${string(reservationFor.arrivalTime)})}}",
                                        "color": "attention",
                                        "horizontalAlignment": "right",
                                        "weight": "bolder",
                                        "spacing": "small",
                                        "wrap": true
                                    }
                                ]
                            }
                        ]
                    },
                    {
                        "type": "ColumnSet",
                        "spacing": "medium",
                        "separator": true,
                        "columns": [
                            {
                                "type": "Column",
                                "width": 1,
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "${reservationFor.departureAirport.name}",
                                        "isSubtle": true,
                                        "wrap": true
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": "${reservationFor.departureAirport.iataCode}",
                                        "size": "extraLarge",
                                        "color": "accent",
                                        "spacing": "none",
                                        "wrap": true
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": "auto",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": " ",
                                        "wrap": true
                                    },
                                    {
                                        "type": "Image",
                                        "url": "https://adaptivecards.io/content/airplane.png",
                                        "altText": "Airplane",
                                        "size": "small"
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": 1,
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "${reservationFor.arrivalAirport.name}",
                                        "isSubtle": true,
                                        "horizontalAlignment": "right",
                                        "wrap": true
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": "${reservationFor.arrivalAirport.iataCode}",
                                        "horizontalAlignment": "right",
                                        "size": "extraLarge",
                                        "color": "accent",
                                        "spacing": "none",
                                        "wrap": true
                                    }
                                ]
                            }
                        ]
                    }
                ]
            }

            const data = {
                "@context": "http://schema.org",
                "@type": "FlightReservation",
                "reservationId": "RXJ34P",
                "reservationStatus": "http://schema.org/ReservationConfirmed",
                "passengerPriorityStatus": "Fast Track",
                "passengerSequenceNumber": "ABC123",
                "securityScreening": "TSA PreCheck",
                "underName": {
                    "@type": "Person",
                    "name": "Sarah Hum"
                },
                "reservationFor": {
                    "@type": "Flight",
                    "flightNumber": "KL605",
                    "provider": {
                        "@type": "Airline",
                        "name": "KLM",
                        "iataCode": "KL",
                        "boardingPolicy": "http://schema.org/ZoneBoardingPolicy"
                    },
                    "seller": {
                        "@type": "Airline",
                        "name": "KLM",
                        "iataCode": "KL"
                    },
                    "departureAirport": {
                        "@type": "Airport",
                        "name": "Amsterdam Airport",
                        "iataCode": "AMS"
                    },
                    "departureTime": "2017-03-04T09:20:00-01:00",
                    "arrivalAirport": {
                        "@type": "Airport",
                        "name": "San Francisco Airport",
                        "iataCode": "SFO"
                    },
                    "arrivalTime": "2017-03-05T08:20:00+04:00"
                }
            }

            const adaptiveCard = AdaptiveCards.declare(template).render(data);
           card = CardFactory.adaptiveCard(adaptiveCard);
        }

        const preview = CardFactory.heroCard(`pill: ${name}`); // pill

        const attachment = { ...card, preview };

        return {
            composeExtension: {
                type: 'result',
                attachmentLayout: 'list',
                attachments: [attachment]
            }
        };
    }

    /* Messaging Extension - Unfurling Link */
    handleTeamsAppBasedLinkQuery(context, query) {
        const attachment = CardFactory.thumbnailCard('Thumbnail Card',
            query.url,
            ['https://raw.githubusercontent.com/microsoft/botframework-sdk/master/icon.png']);

        const result = {
            attachmentLayout: 'list',
            type: 'result',
            attachments: [attachment]
        };

        const response = {
            composeExtension: result
        };
        return response;
    }
    /* Messaging Extension - Unfurling Link */
}

module.exports.BotActivityHandler = BotActivityHandler;

