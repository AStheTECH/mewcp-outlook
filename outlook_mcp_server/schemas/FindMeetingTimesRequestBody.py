"""
Pydantic schemas for the Microsoft Graph `findMeetingTimes` API.

Reference:
https://learn.microsoft.com/en-us/graph/api/user-findmeetingtimes?view=graph-rest-1.0&tabs=http
"""

from enum import Enum
from typing import Annotated, List, Optional

from pydantic import BaseModel, Field


# ---------------------------------------------------------------------------
# Shared building blocks
# ---------------------------------------------------------------------------

class EmailAddress(BaseModel):
    address: Annotated[str, Field(description="Email address of the attendee")]
    name: Annotated[
        Optional[str], Field(default=None, description="Display name of the attendee")
    ] = None


class AttendeeType(str, Enum):
    """The attendeeBase 'type' property. Graph treats every person attendee
    as required by default, so 'required' is used as the schema default."""
    required = "required"
    optional = "optional"
    resource = "resource"


class Attendee(BaseModel):
    """Mirrors the `attendeeBase` resource:
    { "type": "String", "emailAddress": {...} }
    """
    emailAddress: Annotated[
        EmailAddress,
        Field(description="The name and email address (SMTP address) of an attendee"),
    ]
    type: Annotated[
        Optional[AttendeeType],
        Field(
            default=AttendeeType.required,
            description=(
                "The type of attendee: 'required' for a person, "
                "'resource' for a room/equipment, or 'optional'."
            ),
        ),
    ] = AttendeeType.required


class DateTimeZone(BaseModel):
    dateTime: Annotated[
        str,
        Field(
            description=(
                "A single point of time in a combined date and time representation "
                "(for example, 2017-08-29T04:00:00.0000000)."
            )
        ),
    ]
    timeZone: Annotated[
        str,
        Field(default="UTC", description="Represents a time zone (Windows time zone name, e.g. 'Pacific Standard Time')."),
    ] = "UTC"


class TimeSlot(BaseModel):
    start: Annotated[DateTimeZone, Field(description="The date, time, and time zone that a period begins.")]
    end: Annotated[DateTimeZone, Field(description="The date, time, and time zone that a period ends.")]


class ActivityDomain(str, Enum):
    work = "work"
    personal = "personal"
    unrestricted = "unrestricted"
    # 'unknown' is deprecated by Microsoft; intentionally omitted.


class TimeConstraint(BaseModel):
    """Fixed: was previously `class TimeConstraint[BaseModel]:`, which is
    invalid Python syntax and would raise a TypeError on import."""
    activityDomain: Annotated[
        Optional[ActivityDomain],
        Field(
            default=ActivityDomain.work,
            description=(
                "The nature of the meeting, restricting suggested hours. "
                "Defaults to 'work' if not specified."
            ),
        ),
    ] = ActivityDomain.work
    timeSlots: Annotated[
        List[TimeSlot], Field(description="An array of time periods to search within.")
    ]


class Location(BaseModel):
    displayName: Annotated[
        Optional[str], Field(default=None, description="Display name of the location.")
    ] = None
    resolveAvailability: Annotated[
        Optional[bool],
        Field(default=None, description="Whether to resolve availability for this location."),
    ] = None


class LocationConstraint(BaseModel):
    isRequired: Annotated[
        Optional[bool],
        Field(default=False, description="Whether a location suggestion is required."),
    ] = False
    suggestLocation: Annotated[
        Optional[bool],
        Field(default=False, description="Whether findMeetingTimes should suggest a location."),
    ] = False
    locations: Annotated[
        Optional[List[Location]],
        Field(default=None, description="Specific locations the meeting can take place in."),
    ] = None


# ---------------------------------------------------------------------------
# Request body
# ---------------------------------------------------------------------------

class FindMeetingTimesRequestBody(BaseModel):
    """All parameters are Optional per the Graph documentation table —
    there are no strictly required fields for this endpoint."""

    attendees: Annotated[
        Optional[List[Attendee]],
        Field(
            default=None,
            description=(
                "A collection of attendees or resources for the meeting. An empty "
                "collection causes findMeetingTimes to look for free time slots "
                "for only the organizer."
            ),
        ),
    ] = None

    locationConstraint: Annotated[
        Optional[LocationConstraint],
        Field(default=None, description="The organizer's requirements about the meeting location."),
    ] = None

    isOrganizerOptional: Annotated[
        Optional[bool],
        Field(
            default=False,
            description="Specify true if the organizer doesn't necessarily have to attend. Defaults to false.",
        ),
    ] = False

    maxCandidates: Annotated[
        Optional[int],
        Field(default=None, description="The maximum number of meeting time suggestions to be returned."),
    ] = None

    meetingDuration: Annotated[
        Optional[str],
        Field(
            default=None,
            description=(
                "The length of the meeting, in ISO8601 duration format "
                "(e.g. 'PT1H' for 1 hour, 'PT2H30M' for 2.5 hours). "
                "Defaults to 30 minutes if not specified."
            ),
        ),
    ] = None

    # Fixed: was `int`, but Graph documents this as Edm.Double (a % from 0-100).
    minimumAttendeePercentage: Annotated[
        Optional[float],
        Field(
            default=100.0,
            le=100,
            ge=0,
            description="The minimum required confidence (0-100) for a time slot to be returned in the response.",
        ),
    ] = 100.0

    returnSuggestionReasons: Annotated[
        Optional[bool],
        Field(
            default=False,
            description="Specify true to return a reason for each suggestion in suggestionReason. Defaults to false.",
        ),
    ] = False

    timeConstraint: Annotated[
        Optional[TimeConstraint],
        Field(
            default=None,
            description=(
                "Any time restrictions for a meeting, including activityDomain "
                "and possible meeting time periods (timeSlots)."
            ),
        ),
    ] = None
