select
  s.uid,
  s.jur,
  s.house_id,
  s.city,
  s.addr,
  s.tname,
  s.speed,
  s.price
from
  stat_tariffs_addresses as s
where
  s.date = current_date
group by uid
order by addr