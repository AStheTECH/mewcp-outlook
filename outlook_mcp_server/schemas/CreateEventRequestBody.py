from pydantic import BaseModel, Field
from typing import Annotated, List, Optional
from datetime import datetime, timezone

class EmailAddress(BaseModel):
    address: Annotated[str, Field(description="email address of the attendee")]
    name: Annotated[Optional[str], Field(default=None, description="display name of the attendee")] = None

class Attendee(BaseModel):
    emailAddress: Annotated[EmailAddress, Field(description="email address object of the attendee")]
    type: Annotated[str, Field(default="required", description="required or optional attendee")] = "required"

class EventBody(BaseModel):
    contentType: Annotated[str, Field(default="HTML", description="Content Type for the Event Body")] = "HTML"
    content: Annotated[str, Field(description="Body content for the event")]

class EventTime(BaseModel):
    dateTime: Annotated[str, Field(default_factory=lambda: datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S'), description="dateTime in the format YYYY-MM-DDTHH:MM:SS")]
    timeZone: Annotated[str, Field(default="UTC", description="timezone of event time, e.g. 'UTC', 'India Standard Time'")] = "UTC"

class EventLocation(BaseModel):
    displayName: Annotated[str, Field(description="Event Location Name")]

class CreateEventRequestBody(BaseModel):
    subject: Annotated[str, Field(description="Subject Line for the Event")]
    body: Annotated[EventBody, Field(description="body for the event")]
    start: Annotated[EventTime, Field(description="Event start time")]
    end: Annotated[EventTime, Field(description="Event end time")]
    location: Optional[Annotated[EventLocation, Field(description="Location description for the event")]] = None
    attendees: Annotated[List[Attendee], Field(description="List of attendees, with the email address and names")]
