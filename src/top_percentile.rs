use std::cmp::Reverse;
use std::collections::BinaryHeap;

#[derive(Copy, Clone, Debug, PartialEq)]
pub struct F64Ord(pub f64);

impl Eq for F64Ord {}

impl Ord for F64Ord {
    fn cmp(&self, other: &Self) -> std::cmp::Ordering {
        // Explicitly reject NaN keys by panicking.
        if self.0.is_nan() || other.0.is_nan() {
            panic!("NaN keys are not allowed");
        }
        self.0.total_cmp(&other.0)
    }
}
impl PartialOrd for F64Ord {
    fn partial_cmp(&self, other: &Self) -> Option<std::cmp::Ordering> {
        Some(self.cmp(other))
    }
}

#[derive(Debug, Clone)]
struct Entry<K, D> {
    key: K,
    data: D,
    seq: u64,
}

impl<K: Ord, D> Ord for Entry<K, D> {
    fn cmp(&self, other: &Self) -> std::cmp::Ordering {
        // Order by key, then seq (older first) to make a total order.
        self.key
            .cmp(&other.key)
            .then_with(|| self.seq.cmp(&other.seq))
    }
}
impl<K: Ord, D> PartialOrd for Entry<K, D> {
    fn partial_cmp(&self, other: &Self) -> Option<std::cmp::Ordering> {
        Some(self.cmp(other))
    }
}
impl<K: Ord, D> PartialEq for Entry<K, D> {
    fn eq(&self, other: &Self) -> bool {
        self.key == other.key && self.seq == other.seq
    }
}
impl<K: Ord, D> Eq for Entry<K, D> {}

pub struct TopPercentile<K: Ord, D> {
    capacity: usize,
    primary_limit: usize,
    seq: u64,
    primary: BinaryHeap<Reverse<Entry<K, D>>>,
    overflow: BinaryHeap<Reverse<Entry<K, D>>>,
}

impl<K: Ord, D> TopPercentile<K, D> {
    pub fn new(percentage: f64, capacity: usize) -> Self {
        let primary_limit = ((1.0 - percentage) * capacity as f64) as usize;
        if primary_limit == 0 {
            panic!("TopK: primary limit must be greater than 0");
        }
        let overflow_limit = capacity - primary_limit;
        if overflow_limit == 0 {
            panic!("TopK: overflow limit must be greater than 0");
        }
        Self {
            capacity,
            primary_limit,
            seq: 0,
            primary: BinaryHeap::with_capacity(primary_limit),
            overflow: BinaryHeap::with_capacity(overflow_limit),
        }
    }

    pub fn len_primary(&self) -> usize {
        self.primary.len()
    }
    pub fn len_overflow(&self) -> usize {
        self.overflow.len()
    }
    pub fn len_total(&self) -> usize {
        self.primary.len() + self.overflow.len()
    }
    pub fn is_empty(&self) -> bool {
        self.len_total() == 0
    }

    pub fn insert(&mut self, key: K, data: D) {
        let e = Entry {
            key,
            data,
            seq: {
                let s = self.seq;
                self.seq += 1;
                s
            },
        };
        if self.primary.len() < self.primary_limit {
            self.primary.push(Reverse(e));
        } else if let Some(mut smallest) = self.primary.peek_mut() {
            if e > smallest.0 {
                let evicted = std::mem::replace(&mut *smallest, Reverse(e));
                drop(smallest);
                self.overflow.push(evicted);
            } else {
                drop(smallest);
                self.overflow.push(Reverse(e));
            }
        } else {
            self.overflow.push(Reverse(e));
        }
        if self.len_total() > self.capacity {
            panic!("TopK: total entries exceeded capacity");
        }
    }

    pub fn threshold(&self) -> Option<&K> {
        self.primary.peek().map(|r| &r.0.key)
    }

    pub fn smallest_primary(&self) -> Option<(&K, &D)> {
        self.primary.peek().map(|r| (&r.0.key, &r.0.data))
    }

    pub fn iter_primary_desc(&self) -> impl Iterator<Item = (&K, &D)> {
        let mut refs: Vec<&Entry<K, D>> = self.primary.iter().map(|r| &r.0).collect();
        refs.sort_by(|a, b| b.key.cmp(&a.key).then_with(|| b.seq.cmp(&a.seq)));
        refs.into_iter().map(|e| (&e.key, &e.data))
    }

    pub fn iter_overflow_desc(&self) -> impl Iterator<Item = (&K, &D)> {
        let mut refs: Vec<&Entry<K, D>> = self.overflow.iter().map(|r| &r.0).collect();
        refs.sort_by(|a, b| b.key.cmp(&a.key).then_with(|| b.seq.cmp(&a.seq)));
        refs.into_iter().map(|e| (&e.key, &e.data))
    }
}
